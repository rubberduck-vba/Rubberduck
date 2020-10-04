using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using IOException = System.IO.IOException;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using NLog;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
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

        private readonly List<string> _unloadedCOMReferences = new List<string>();
        private readonly List<(string projectId, string referencedProjectId)> _referencesAffectedByPriorityChanges = new List<(string projectId, string referencedProjectId)>();

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
        }


        public bool LastSyncOfCOMReferencesLoadedReferences { get; private set; }
        public IEnumerable<string> COMReferencesUnloadedInLastSync => _unloadedCOMReferences;
        public IEnumerable<(string projectId, string referencedProjectId)> COMReferencesAffectedByPriorityChangesInLastSync => _referencesAffectedByPriorityChanges;

        private readonly IDictionary<string, ReferencePriorityMap> _projectReferences = new Dictionary<string, ReferencePriorityMap>();
        public IReadOnlyCollection<ReferencePriorityMap> ProjectReferences => _projectReferences.Values.AsReadOnly();


        protected abstract void LoadReferences(IEnumerable<ReferenceInfo> referencesToLoad, ConcurrentBag<ReferenceInfo> unmapped, CancellationToken token);


        public void SyncComReferences(CancellationToken token)
        {
            var parsingStageTimer = ParsingStageTimer.StartNew();

            var oldProjectReferences = _projectReferences.ToDictionary();
            _projectReferences.Clear();

            LastSyncOfCOMReferencesLoadedReferences = false;
            _unloadedCOMReferences.Clear();
            _referencesAffectedByPriorityChanges.Clear();
            RefreshReferencedByProjectId();

            var unmapped = new ConcurrentBag<ReferenceInfo>();

            var referencesByProjectId = ReferencedByProjectId();
            var referencesToLoad = GetReferencesToLoadAndSaveReferencePriority(referencesByProjectId, oldProjectReferences);

            if (referencesToLoad.Any())
            {
                LastSyncOfCOMReferencesLoadedReferences = true;
                LoadReferences(referencesToLoad, unmapped, token);
            }

            DetermineReferencesAffectedByPriorityChanges(_projectReferences, oldProjectReferences);
            UnloadNoLongerExistingReferences(_projectReferences, oldProjectReferences);

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Loaded and unloaded referenced libraries in {0}ms.");
        }

        private void UnloadNoLongerExistingReferences(IDictionary<string, ReferencePriorityMap> newProjectReferences, IDictionary<string, ReferencePriorityMap> oldProjectReferences)
        {
            var noLongerReferencedProjectIds = oldProjectReferences.Keys
                .Where(projectId => !newProjectReferences.ContainsKey(projectId))
                .ToList();

            foreach (var referencedProjectId in noLongerReferencedProjectIds)
            {
                UnloadComReference(referencedProjectId);
            }
        }

        private void DetermineReferencesAffectedByPriorityChanges(IDictionary<string, ReferencePriorityMap> newProjectReferences, IDictionary<string, ReferencePriorityMap> oldProjectReferences)
        {
            var referencePriorityChanges = ReferencePriorityChanges(newProjectReferences, oldProjectReferences);

            foreach (var oldMap in oldProjectReferences.Values)
            {
                foreach (var projectId in referencePriorityChanges.Keys)
                {
                    if (oldMap.TryGetValue(projectId, out var priority)
                        && referencePriorityChanges[projectId]
                            .Any(tpl => (tpl.oldPriority <= priority && tpl.newPriority >= priority)
                                        || (tpl.oldPriority >= priority && tpl.newPriority <= priority)))
                    {
                        _referencesAffectedByPriorityChanges.Add((projectId, oldMap.ReferencedProjectId));
                    }
                }
            }
        }

        private static IDictionary<string, List<(int newPriority, int oldPriority)>> ReferencePriorityChanges(IDictionary<string, ReferencePriorityMap> newProjectReferences, IDictionary<string, ReferencePriorityMap> oldProjectReferences)
        {
            var referencePriorityChanges = new Dictionary<string, List<(int newPriority, int oldPriority)>>();
            foreach (var referencedProjectId in oldProjectReferences.Keys)
            {
                var oldMap = oldProjectReferences[referencedProjectId];
                if (!newProjectReferences.TryGetValue(referencedProjectId, out var newMap))
                {
                    foreach (var projectId in oldMap.Keys)
                    {
                        AddPriorityChangeToDictionary(projectId, int.MaxValue, oldMap[projectId], referencePriorityChanges);
                    }
                }
                else
                {
                    foreach (var projectId in oldMap.Keys)
                    {
                        var oldPriority = oldMap[projectId];
                        if (!newMap.TryGetValue(referencedProjectId, out var newPriority))
                        {
                            newPriority = int.MaxValue;
                        }

                        if (newPriority != oldPriority)
                        {
                            AddPriorityChangeToDictionary(projectId, newPriority, oldMap[projectId], referencePriorityChanges);
                        }
                    }

                    foreach (var projectId in newMap.Keys)
                    {
                        if (!oldMap.ContainsKey(projectId))
                        {
                            AddPriorityChangeToDictionary(projectId, newMap[projectId], int.MaxValue, referencePriorityChanges);
                        }
                    }
                }
            }

            foreach (var referencedProjectId in newProjectReferences.Keys)
            {
                if (oldProjectReferences.ContainsKey(referencedProjectId))
                {
                    continue;
                }

                var newMap = newProjectReferences[referencedProjectId];
                foreach (var projectId in newMap.Keys)
                {
                    AddPriorityChangeToDictionary(projectId, newMap[projectId], int.MaxValue, referencePriorityChanges);
                }
            }

            return referencePriorityChanges;
        }

        private static void AddPriorityChangeToDictionary(
            string projectId, 
            int newPriority, 
            int oldPriority,
            IDictionary<string, List<(int newPriority, int oldPriority)>> dict)
        {
            if (dict.TryGetValue(projectId, out var changeList))
            {
                changeList.Add((newPriority, oldPriority));
            }
            else
            {
                dict.Add(projectId, new List<(int newPriority, int oldPriority)> { (newPriority, oldPriority) });
            }
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
            var lockedProjects = _projectsProvider.LockedProjects();
            foreach (var (projectId, project) in projects.Concat(lockedProjects))
            {
                if (TryGetFullPath(project, out var fullPath))
                {
                    var projectName = !string.IsNullOrEmpty(fullPath) ? project.Name : $"UnsavedProject{projectId}";
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

        private ICollection<ReferenceInfo> GetReferencesToLoadAndSaveReferencePriority(Dictionary<string, List<ReferenceInfo>> referencedByProjectId, IDictionary<string, ReferencePriorityMap> oldProjectReferences)
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
                    if (!_projectReferences.TryGetValue(referencedProjectId, out var map))
                    {
                        map = new ReferencePriorityMap(referencedProjectId) { { projectId, priority } };
                        _projectReferences.Add(referencedProjectId, map);

                        if (oldProjectReferences.TryGetValue(referencedProjectId, out var oldMap))
                        {
                            map.IsLoaded = oldMap.IsLoaded;
                        }
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
            return references.Where(item => !_projectReferences.ContainsKey(GetReferenceProjectId(item))).ToList();
        }

        private void UnloadComReference(string referencedProjectId)
        {
            //There is nothing to unload for a user project.
            if (!IsUserProjectProjectId(referencedProjectId))
            {
                _unloadedCOMReferences.Add(referencedProjectId);

                _state.RemoveBuiltInDeclarations(referencedProjectId);
            }
        }

        private bool IsUserProjectProjectId(string projectId)
        {
            return _projectIdsByFilePathAndProjectName.Values.Contains(projectId);
        }
    }
}
