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
        private readonly IModuleToModuleReferenceManager _moduleToModuleReferenceManager;

        protected ReferenceRemoverBase(
            RubberduckParserState state,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager)
        {
            _state = state ?? throw new ArgumentNullException(nameof(state));

            _moduleToModuleReferenceManager = moduleToModuleReferenceManager ?? throw new ArgumentNullException(nameof(moduleToModuleReferenceManager));
        }

        public abstract void RemoveReferencesTo(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
        protected abstract void RemoveReferencesByFromTargetModules(IReadOnlyCollection<QualifiedModuleName> referencingModules, IReadOnlyCollection<QualifiedModuleName> targetModules, CancellationToken token);


        public void RemoveReferencesBy(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token)
        {
            if (!modules.Any())
            {
                return;
            }
            var modulesNeedingReferenceRemoval = _moduleToModuleReferenceManager.ModulesReferencedByAny(modules);
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
