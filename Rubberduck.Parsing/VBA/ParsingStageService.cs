using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public class ParsingStageService : IParsingStageService
    {
        private readonly ICOMReferenceSynchronizer _comSynchronizer;
        private readonly IBuiltInDeclarationLoader _builtInDeclarationLoader;
        private readonly IParseRunner _parseRunner;
        private readonly IDeclarationResolveRunner _declarationResolver;
        private readonly IReferenceResolveRunner _referenceResolver;
        private readonly IUserComProjectSynchronizer _userComProjectSynchronizer;

        public ParsingStageService(
            ICOMReferenceSynchronizer comSynchronizer,
            IBuiltInDeclarationLoader builtInDeclarationLoader,
            IParseRunner parseRunner,
            IDeclarationResolveRunner declarationResolver,
            IReferenceResolveRunner referenceResolver,
            IUserComProjectSynchronizer userComProjectSynchronizer)
        {
            _comSynchronizer = comSynchronizer ?? throw new ArgumentNullException(nameof(comSynchronizer));
            _builtInDeclarationLoader = builtInDeclarationLoader ?? throw new ArgumentNullException(nameof(builtInDeclarationLoader));
            _parseRunner = parseRunner ?? throw new ArgumentNullException(nameof(parseRunner));
            _declarationResolver = declarationResolver ?? throw new ArgumentNullException(nameof(declarationResolver));
            _referenceResolver = referenceResolver ?? throw new ArgumentNullException(nameof(referenceResolver));
            _userComProjectSynchronizer = userComProjectSynchronizer ?? throw new ArgumentNullException(nameof(userComProjectSynchronizer));
        }

        public bool LastLoadOfBuiltInDeclarationsLoadedDeclarations => _builtInDeclarationLoader.LastLoadOfBuiltInDeclarationsLoadedDeclarations;
        public bool LastSyncOfCOMReferencesLoadedReferences => _comSynchronizer.LastSyncOfCOMReferencesLoadedReferences;
        public IEnumerable<string> COMReferencesUnloadedInLastSync => _comSynchronizer.COMReferencesUnloadedInLastSync;
        public IEnumerable<(string projectId, string referencedProjectId)>COMReferencesAffectedByPriorityChangesInLastSync =>_comSynchronizer.COMReferencesAffectedByPriorityChangesInLastSync;

        public void LoadBuitInDeclarations()
        {
            _builtInDeclarationLoader.LoadBuitInDeclarations();
        }

        public void ParseModules(IReadOnlyCollection<QualifiedModuleName> modulesToParse, CancellationToken token)
        {
            _parseRunner.ParseModules(modulesToParse, token);
        }

        public void CreateProjectDeclarations(IReadOnlyCollection<string> projectIds)
        {
            _declarationResolver.CreateProjectDeclarations(projectIds);
        }

        public void RefreshProjectReferences()
        {
            _declarationResolver.RefreshProjectReferences();
        }

        public void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            _declarationResolver.ResolveDeclarations(modules, token);
        }

        public void ResolveReferences(IReadOnlyCollection<QualifiedModuleName> toResolve, CancellationToken token)
        {
            _referenceResolver.ResolveReferences(toResolve, token);
        }

        public void SyncComReferences(CancellationToken token)
        {
            _comSynchronizer.SyncComReferences(token);
        }

        public bool LastSyncOfUserComProjectsLoadedDeclarations =>
            _userComProjectSynchronizer.LastSyncOfUserComProjectsLoadedDeclarations;

        public IReadOnlyCollection<string> UserProjectIdsUnloaded => _userComProjectSynchronizer.UserProjectIdsUnloaded;

        public void SyncUserComProjects()
        {
            _userComProjectSynchronizer.SyncUserComProjects();
        }
    }
}
