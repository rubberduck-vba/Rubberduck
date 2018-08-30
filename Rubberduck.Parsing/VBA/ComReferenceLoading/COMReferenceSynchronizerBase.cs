using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using NLog;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public abstract class COMReferenceSynchronizerBase : ICOMReferenceSynchronizer, IProjectReferencesProvider 
    {
        protected readonly RubberduckParserState _state;
        protected readonly IParserStateManager _parserStateManager;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IReferencedDeclarationsCollector _referencedDeclarationsCollector;

        private readonly List<QualifiedModuleName> _unloadedCOMReferences;

        private readonly Dictionary<(string identifierName, string fullPath), string> _projectIdsByFilePathAndProjectName = new Dictionary<(string identifierName, string fullPath), string>();

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();


        protected COMReferenceSynchronizerBase(RubberduckParserState state, IParserStateManager parserStateManager, IProjectsProvider projectsProvider, IReferencedDeclarationsCollector referencedDeclarationsCollector)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (parserStateManager == null)
            {
                throw new ArgumentNullException(nameof(parserStateManager));
            }
            if (projectsProvider == null)
            {
                throw new ArgumentNullException(nameof(projectsProvider));
            }
            if (referencedDeclarationsCollector == null)
            {
                throw new ArgumentNullException(nameof(referencedDeclarationsCollector));
            }

            _state = state;
            _parserStateManager = parserStateManager;
            _projectsProvider = projectsProvider;
            _referencedDeclarationsCollector = referencedDeclarationsCollector;
            _unloadedCOMReferences = new List<QualifiedModuleName>();
        }


        public bool LastSyncOfCOMReferencesLoadedReferences { get; private set; }
        public IEnumerable<QualifiedModuleName> COMReferencesUnloadedUnloadedInLastSync => _unloadedCOMReferences;

        private readonly HashSet<ReferencePriorityMap> _projectReferences = new HashSet<ReferencePriorityMap>();
        public IReadOnlyCollection<ReferencePriorityMap> ProjectReferences => _projectReferences.AsReadOnly();


        protected abstract void LoadReferences(IEnumerable<ReferenceInfo> referencesToLoad, ConcurrentBag<ReferenceInfo> unmapped, CancellationToken token);


        public void SyncComReferences(CancellationToken token)
        {
            var parsingStageTimer = ParsingStageTimer.StartNew();

            LastSyncOfCOMReferencesLoadedReferences = false;
            _unloadedCOMReferences.Clear();
            RefreshReferencedByProjectId();

            var unmapped = new ConcurrentBag<ReferenceInfo>();

            var referencesByProjectId = ReferencedByProjectId();
            var referencesToLoad = GetReferencesToLoadAndSaveReferencePriority(referencesByProjectId);

            if (referencesToLoad.Any())
            {
                LastSyncOfCOMReferencesLoadedReferences = true;
                LoadReferences(referencesToLoad, unmapped, token);
            }

            var allReferences = referencesByProjectId.Values.SelectMany(references => references).ToHashSet();
            var notMappedReferences = NonMappedReferences(allReferences);
            foreach (var item in notMappedReferences)
            {
                unmapped.Add(item);
            }

            foreach (var reference in unmapped)
            {
                UnloadComReference(reference);
            }

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Loaded and unloaded referenced libraries in {0}ms.");
        }

        private Dictionary<string, List<ReferenceInfo>> ReferencedByProjectId()
        {
            var referencesByProjectId = new Dictionary<string, List<ReferenceInfo>>();

            var projects = _projectsProvider.Projects();
            foreach (var (projectId, project) in projects)
            {
                referencesByProjectId.Add(projectId, ProjectReferenceInfos(project));
            }

            return referencesByProjectId;
        }

        private List<ReferenceInfo> ProjectReferenceInfos(IVBProject project)
        {
            var referenceInfos = new List<ReferenceInfo>();
            using (var references = project.References)
            {
                for (var priority = 1; priority <= references.Count; priority++)
                {
                    using (var reference = references[priority])
                    {
                        if (reference.IsBroken)
                        {
                            continue;
                        }
                        referenceInfos.Add(new ReferenceInfo(reference));
                    }
                }
            }

            return referenceInfos;
        }

        private void RefreshReferencedByProjectId()
        {
            _projectIdsByFilePathAndProjectName.Clear();

            var projects = _projectsProvider.Projects();
            foreach (var (projectId, project) in projects)
            {
                if (TryGetFullPath(project, out var fullPath))
                {
                    var projectName = project.Name;
                    _projectIdsByFilePathAndProjectName.Add((projectName, fullPath), projectId);
                }
            }
        }

        private static bool TryGetFullPath(IVBProject project, out string fullPath)
        {
            try
            {
                fullPath = project.FileName;
            }
            catch (IOException)
            {
                // Filename throws exception if unsaved.
                fullPath = null;
                return false;
            }
            catch (COMException e)
            {
                Logger.Warn(e);
                fullPath = null;
                return false;
            }

            return true;
        }

        private ICollection<ReferenceInfo> GetReferencesToLoadAndSaveReferencePriority(Dictionary<string, List<ReferenceInfo>> referencedByProjectId)
        {
            var referencesToLoad = new List<ReferenceInfo>();

            foreach (var (projectId, references) in referencedByProjectId)
            {
                // use a 'for' loop to store the order of references as a 'priority', which is 1-based by VBA convention.
                // reference resolver needs this to know which declaration to prioritize when a global identifier exists in multiple libraries.
                for (var priority = 1; priority <= references.Count; priority++)
                {
                    var reference = references[priority - 1];

                    // todo: figure out why Rubberduck.tlb *sometimes* throws

                    var referencedProjectId = GetReferenceProjectId(reference);
                    var map = _projectReferences.FirstOrDefault(item =>
                        item.ReferencedProjectId == referencedProjectId);

                    if (map == null)
                    {
                        map = new ReferencePriorityMap(referencedProjectId) { { projectId, priority } };
                        _projectReferences.Add(map);
                    }
                    else
                    {
                        map[projectId] = priority;
                    }

                    if (!map.IsLoaded)
                    {
                        //There is nothing to load for a user project.
                        if (!IsUserProjectProjectId(referencedProjectId))
                        {
                            referencesToLoad.Add(reference);
                        }
                        map.IsLoaded = true;
                    }
                }
            }

            return referencesToLoad;
        }

        private string GetReferenceProjectId(ReferenceInfo reference)
        {
            if (_projectIdsByFilePathAndProjectName.TryGetValue((reference.Name, reference.FullPath), out var projectId))
            {
                return projectId;
            }

            return QualifiedModuleName.GetProjectId(reference);
        }

        protected void LoadReference(ReferenceInfo reference, ConcurrentBag<ReferenceInfo> unmapped)
        {
            if (Thread.CurrentThread.IsBackground && Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = $"LoadReference '{reference.Name}'";
            }

            Logger.Trace($"Loading referenced type '{reference.Name}'.");            
            try
            {
                var declarations = _referencedDeclarationsCollector.CollectedDeclarations(reference);
                foreach (var declaration in declarations)
                {
                    _state.AddDeclaration(declaration);
                }
            }
            catch (Exception exception)
            {
                unmapped.Add(reference);
                Logger.Warn($"Types were not loaded from referenced type library '{reference.Name}'.");
                Logger.Error(exception);
            }
        }

        private IEnumerable<ReferenceInfo> NonMappedReferences(ICollection<ReferenceInfo> references)
        {
            var mappedIds = _projectReferences.Select(item => item.ReferencedProjectId).ToHashSet();
            return references.Where(item => !mappedIds.Contains(GetReferenceProjectId(item))).ToList();
        }

        private void UnloadComReference(ReferenceInfo reference)
        {
            var referencedProjectId = GetReferenceProjectId(reference);

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
            if (map.Count == 0)
            {
                _projectReferences.Remove(map);

                //There is nothing to unload for a user project.
                if (!IsUserProjectProjectId(referencedProjectId))
                {
                    AddUnloadedReferenceToUnloadedReferences(reference);
                    _state.RemoveBuiltInDeclarations(reference);
                }
            }
        }

        private bool IsUserProjectProjectId(string projectId)
        {
            return _projectIdsByFilePathAndProjectName.Values.Contains(projectId);
        }

        private void AddUnloadedReferenceToUnloadedReferences(ReferenceInfo reference)
        {
            var projectQMN = new QualifiedModuleName(reference);
            _unloadedCOMReferences.Add(projectQMN);
        }
    }
}
