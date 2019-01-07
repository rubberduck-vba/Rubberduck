using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public abstract class ReferenceRemoverBase : IReferenceRemover
    {
        private readonly RubberduckParserState _state;
        private readonly IModuleToModuleReferenceManager _moduleToModuleReferenceManager;

        public ReferenceRemoverBase(
            RubberduckParserState state,
            IModuleToModuleReferenceManager moduleToModuleReferenceManager)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }
            if (moduleToModuleReferenceManager == null)
            {
                throw new ArgumentNullException(nameof(moduleToModuleReferenceManager));
            }

            _state = state;
            _moduleToModuleReferenceManager = moduleToModuleReferenceManager;
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
