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
            if(comSynchronizer == null)
            {
                throw new ArgumentNullException(nameof(comSynchronizer));
            }
            if (builtInDeclarationLoader == null)
            {
                throw new ArgumentNullException(nameof(builtInDeclarationLoader));
            }
            if (parseRunner == null)
            {
                throw new ArgumentNullException(nameof(parseRunner));
            }
            if (declarationResolver == null)
            {
                throw new ArgumentNullException(nameof(declarationResolver));
            }
            if (referenceResolver == null)
            {
                throw new ArgumentNullException(nameof(referenceResolver));
            }

            _comSynchronizer = comSynchronizer;
            _builtInDeclarationLoader = builtInDeclarationLoader;
            _parseRunner = parseRunner;
            _declarationResolver = declarationResolver;
            _referenceResolver = referenceResolver;
        }


        public bool LastLoadOfBuiltInDeclarationsLoadedDeclarations
        {
            get
            {
                return _builtInDeclarationLoader.LastLoadOfBuiltInDeclarationsLoadedDeclarations;
            }
        }

        public bool LastSyncOfCOMReferencesLoadedReferences
        {
            get
            {
                return _comSynchronizer.LastSyncOfCOMReferencesLoadedReferences;
            }
        }

        public bool LastSyncOfCOMReferencesUnloadedReferences
        {
            get
            {
                return _comSynchronizer.LastSyncOfCOMReferencesUnloadedReferences;
            }
        }

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
