using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using NLog;
using System.Collections.Concurrent;
using Rubberduck.VBEditor;
using System.Runtime.InteropServices;
using System.IO;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.Parsing.VBA
{
    public abstract class COMReferenceSynchronizerBase : ICOMReferenceSynchronizer, IProjectReferencesProvider 
    {
        private const string rubberduckGUID = "{E07C841C-14B4-4890-83E9-8C80B06DD59D}";

        protected readonly RubberduckParserState _state;
        protected readonly IParserStateManager _parserStateManager;
        private readonly string _serializedDeclarationsPath;
        private readonly List<QualifiedModuleName> _unloadedCOMReferences;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();


        protected COMReferenceSynchronizerBase(RubberduckParserState state, IParserStateManager parserStateManager, string serializedDeclarationsPath = null)
        {
            _state = state ?? throw new ArgumentNullException(nameof(state));

            _parserStateManager = parserStateManager ?? throw new ArgumentNullException(nameof(parserStateManager));

            _serializedDeclarationsPath = serializedDeclarationsPath
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "declarations");

            _unloadedCOMReferences = new List<QualifiedModuleName>();
        }


        public bool LastSyncOfCOMReferencesLoadedReferences { get; private set; }
        public IEnumerable<QualifiedModuleName> COMReferencesUnloadedUnloadedInLastSync => _unloadedCOMReferences;

        private readonly HashSet<ReferencePriorityMap> _projectReferences = new HashSet<ReferencePriorityMap>();
        public IReadOnlyCollection<ReferencePriorityMap> ProjectReferences => _projectReferences.ToHashSet().AsReadOnly();


        protected abstract void LoadReferences(IEnumerable<IReference> referencesToLoad, ConcurrentBag<IReference> unmapped, CancellationToken token);


        public void SyncComReferences(IReadOnlyList<IVBProject> projects, CancellationToken token)
        {
            LastSyncOfCOMReferencesLoadedReferences = false;
            _unloadedCOMReferences.Clear();

            var unmapped = new ConcurrentBag<IReference>();

            var referencesToLoad = GetReferencesToLoadAndSaveReferencePriority(projects);

            if (referencesToLoad.Any())
            {
                LastSyncOfCOMReferencesLoadedReferences = true;
                LoadReferences(referencesToLoad, unmapped, token);
            }

            var notMappedReferences = NonMappedReferences(projects);
            foreach (var item in notMappedReferences)
            {
                unmapped.Add(item);
            }

            if (!unmapped.Any())
            {
                return;
            }

            foreach (var reference in unmapped)
            {
                UnloadComReference(reference, projects);
            }
        }

        private IEnumerable<IReference> GetReferencesToLoadAndSaveReferencePriority(IReadOnlyList<IVBProject> projects)
        {
            var referencesToLoad = new List<IReference>();

            foreach (var vbProject in projects)
            {
                var projectId = QualifiedModuleName.GetProjectId(vbProject);
                var references = vbProject.References;

                // use a 'for' loop to store the order of references as a 'priority'.
                // reference resolver needs this to know which declaration to prioritize when a global identifier exists in multiple libraries.
                for (var priority = 1; priority <= references.Count; priority++)
                {
                    var reference = references[priority];
                    if (reference.IsBroken)
                    {
                        continue;
                    }

                    // skip loading Rubberduck.tlb (GUID is defined in AssemblyInfo.cs)
                    if (reference.Guid == rubberduckGUID)
                    {
                        // todo: figure out why Rubberduck.tlb *sometimes* throws
                        //continue;
                    }

                    var referencedProjectId = GetReferenceProjectId(reference, projects);
                    var map = _projectReferences.FirstOrDefault(item => item.ReferencedProjectId == referencedProjectId);

                    if (map == null)
                    {
                        map = new ReferencePriorityMap(referencedProjectId) { { projectId, priority } };
                        _projectReferences.Add(map);
                    }
                    else
                    {
                        map[projectId] = priority;
                    }

                    if (map.IsLoaded)
                    {
                        continue;
                    }

                    referencesToLoad.Add(reference);
                    map.IsLoaded = true;
                }
            }
            return referencesToLoad;
        }

        private string GetReferenceProjectId(IReference reference, IReadOnlyList<IVBProject> projects)
        {
            IVBProject project = null;
            foreach (var item in projects)
            {
                try
                {
                    // check the name not just the path, because path is empty in tests:
                    if (item.Name == reference.Name && item.FileName == reference.FullPath)
                    {
                        project = item;
                        break;
                    }
                }
                catch (IOException)
                {
                    // Filename throws exception if unsaved.
                }
                catch (COMException e)
                {
                    Logger.Warn(e);
                }
            }

            if (project == null)
            {
                return QualifiedModuleName.GetProjectId(reference);
            }

            if (string.IsNullOrEmpty(project.ProjectId))
            {
                project.AssignProjectId();
            }
            return project.ProjectId;
        }

        protected void LoadReference(IReference localReference, ConcurrentBag<IReference> unmapped)
        {
            Logger.Trace($"Loading referenced type '{localReference.Name}'.");
            var comReflector = new ReferencedDeclarationsCollector(_state, localReference, _serializedDeclarationsPath);
            try
            {
                if (comReflector.SerializedVersionExists)
                {
                    LoadReferenceByDeserialization(localReference, comReflector);
                }
                else
                {
                    LoadReferenceFromTypeLibrary(localReference, comReflector);
                }
            }
            catch (Exception exception)
            {
                unmapped.Add(localReference);
                Logger.Warn($"Types were not loaded from referenced type library '{localReference.Name}'.");
                Logger.Error(exception);
            }
        }

        private void LoadReferenceByDeserialization(IReference localReference, ReferencedDeclarationsCollector comReflector)
        {
            Logger.Trace($"Deserializing reference '{localReference.Name}'.");
            var declarations = comReflector.LoadDeclarationsFromXml();
            foreach (var declaration in declarations)
            {
                _state.AddDeclaration(declaration);
            }
        }

        private void LoadReferenceFromTypeLibrary(IReference localReference, ReferencedDeclarationsCollector comReflector)
        {
            Logger.Trace($"COM reflecting reference '{localReference.Name}'.");
            var declarations = comReflector.LoadDeclarationsFromLibrary();
            foreach (var declaration in declarations)
            {
                _state.AddDeclaration(declaration);
            }
        }

        private IEnumerable<IReference> NonMappedReferences(IReadOnlyList<IVBProject> projects)
        {
            var mappedIds = _projectReferences.Select(item => item.ReferencedProjectId).ToHashSet();
            var references = projects.SelectMany(project => project.References);
            return references.Where(item => !mappedIds.Contains(GetReferenceProjectId(item, projects))).ToList();
        }

        private void UnloadComReference(IReference reference, IReadOnlyList<IVBProject> projects)
        {
            var referencedProjectId = GetReferenceProjectId(reference, projects);

            ReferencePriorityMap map = null;
            try
            {
                map = _projectReferences.SingleOrDefault(item => item.ReferencedProjectId == referencedProjectId);
            }
            catch (InvalidOperationException exception)
            {
                //There are multiple maps with the same referencedProjectId. That should not happen. (ghost?).
                Logger.Error(exception, "Failed To unload com reference with referencedProjectID {0} because RD stores multiple instances of it.", referencedProjectId);
                return;
            }

            if (map == null || !map.IsLoaded)
            {
                Logger.Warn("Tried to unload untracked project reference."); //This shouldn't happen.
                return;
            }

            map.Remove(referencedProjectId);
            if (map.Count != 0)
            {
                return;
            }

            AddUnloadedReferenceToUnloadedReferences(reference);
            _projectReferences.Remove(map);
            _state.RemoveBuiltInDeclarations(reference);
        }

        private void AddUnloadedReferenceToUnloadedReferences(IReference reference)
        {
            var projectName = reference.Name;
            var projectQMN = new QualifiedModuleName(projectName, reference.FullPath, projectName);
            _unloadedCOMReferences.Add(projectQMN);
        }

    }
}
