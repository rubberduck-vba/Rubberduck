using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;
using System.Threading;
using System.Linq;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ReferenceRemoverBase : IReferenceRemover
    {
        private readonly RubberduckParserState _state;

        public ReferenceRemoverBase(
            RubberduckParserState state)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            _state = state;
        }


        public abstract void RemoveReferencesTo(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
        protected abstract void RemoveReferencesByFromTargetModules(IReadOnlyCollection<QualifiedModuleName> referencingModules, IReadOnlyCollection<QualifiedModuleName> targetModules, CancellationToken token);


        public void RemoveReferencesBy(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }
            var modulesNeedingReferenceRemoval = _state.DeclarationFinder.AllModules();
            RemoveReferencesByFromTargetModules(modules, modulesNeedingReferenceRemoval, token);
        }

        protected void RemoveReferencesByFromTargetModule(IReadOnlyCollection<QualifiedModuleName> referencingModules, QualifiedModuleName targetModule)
        {
            var declarationsInTargetModule = _state.DeclarationFinder.Members(targetModule);
            foreach (var declaration in declarationsInTargetModule)
            {
                declaration.RemoveReferencesFrom(referencingModules);
            }
        }

        public void RemoveReferencesBy(QualifiedModuleName module, CancellationToken token)
        {
            var modules = new HashSet<QualifiedModuleName> { module }.AsReadOnly();
            RemoveReferencesBy(modules, token);
        }

        public void RemoveReferencesTo(QualifiedModuleName module, CancellationToken token)
        {
            var declarationsToClearOfReferences = _state.DeclarationFinder.Members(module);
            foreach(var declaration in declarationsToClearOfReferences)
            {
                declaration.ClearReferences();
            }
        }
    }
}
