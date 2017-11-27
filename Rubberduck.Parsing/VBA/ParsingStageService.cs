using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class ParsingStageService : IParsingStageService
    {
        private readonly ICOMReferenceSynchronizer _comSynchronizer;
        private readonly IBuiltInDeclarationLoader _builtInDeclarationLoader;
        private readonly IParseRunner _parseRunner;
        private readonly IDeclarationResolveRunner _declarationResolver;
        private readonly IReferenceResolveRunner _referenceResolver;

        public ParsingStageService(
            ICOMReferenceSynchronizer comSynchronizer,
            IBuiltInDeclarationLoader builtInDeclarationLoader,
            IParseRunner parseRunner,
            IDeclarationResolveRunner declarationResolver,
            IReferenceResolveRunner referenceResolver)
        {
            _comSynchronizer = comSynchronizer ?? throw new ArgumentNullException(nameof(comSynchronizer));

            _builtInDeclarationLoader = builtInDeclarationLoader ?? throw new ArgumentNullException(nameof(builtInDeclarationLoader));

            _parseRunner = parseRunner ?? throw new ArgumentNullException(nameof(parseRunner));

            _declarationResolver = declarationResolver ?? throw new ArgumentNullException(nameof(declarationResolver));

            _referenceResolver = referenceResolver ?? throw new ArgumentNullException(nameof(referenceResolver));
        }


        public bool LastLoadOfBuiltInDeclarationsLoadedDeclarations => _builtInDeclarationLoader.LastLoadOfBuiltInDeclarationsLoadedDeclarations;

        public bool LastSyncOfCOMReferencesLoadedReferences => _comSynchronizer.LastSyncOfCOMReferencesLoadedReferences;

        public IEnumerable<QualifiedModuleName> COMReferencesUnloadedUnloadedInLastSync => _comSynchronizer.COMReferencesUnloadedUnloadedInLastSync;

        public void LoadBuitInDeclarations()
        {
            _builtInDeclarationLoader.LoadBuitInDeclarations();
        }

        public void ParseModules(IReadOnlyCollection<QualifiedModuleName> modulesToParse, CancellationToken token)
        {
            _parseRunner.ParseModules(modulesToParse, token);
        }

        public void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            _declarationResolver.ResolveDeclarations(modules, token);
        }

        public void ResolveReferences(IReadOnlyCollection<QualifiedModuleName> toResolve, CancellationToken token)
        {
            _referenceResolver.ResolveReferences(toResolve, token);
        }

        public void SyncComReferences(IReadOnlyList<IVBProject> projects, CancellationToken token)
        {
            _comSynchronizer.SyncComReferences(projects, token);
        }
    }
}
