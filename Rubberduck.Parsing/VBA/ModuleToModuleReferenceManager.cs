using System.Collections.Generic;
using Rubberduck.VBEditor;
using System.Collections.Concurrent;

namespace Rubberduck.Parsing.VBA
{
    public class ModuleToModuleReferenceManager : ModuleToModuleReferenceManagerBase
    {
        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>> _referencesFrom = new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>>();
        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>> _referencesTo = new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>>();

        public override void AddModuleToModuleReference(QualifiedModuleName referencingModule, QualifiedModuleName referencedModule)
        {
            ConcurrentDictionary<QualifiedModuleName, byte> referencedModules;
            ConcurrentDictionary<QualifiedModuleName, byte> referencingModules;
            while(!_referencesFrom.TryGetValue(referencingModule, out referencedModules) || !_referencesTo.TryGetValue(referencedModule, out referencingModules))
            {
                _referencesFrom.AddOrUpdate(referencingModule,
                    new ConcurrentDictionary<QualifiedModuleName, byte>(),
                    (key, value) => value);
                _referencesTo.AddOrUpdate(referencedModule,
                    new ConcurrentDictionary<QualifiedModuleName, byte>(),
                    (key, value) => value);
            }
            referencedModules.AddOrUpdate(referencedModule, 0, (key, value) => value);
            referencingModules.AddOrUpdate(referencingModule, 0, (key, value) => value);
        }

        public override void RemoveModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule)
        {
            if (_referencesFrom.TryGetValue(referencingModule, out var referencedModules))
            {
                referencedModules.TryRemove(referencedModule, out var _);
            }

            if (_referencesTo.TryGetValue(referencedModule, out var referencingModules))
            {
                referencingModules.TryRemove(referencingModule, out var _);
            }
        }

        public override void ClearModuleToModuleReferencesToModule(QualifiedModuleName referencedModule)
        {
            if (!_referencesTo.TryRemove(referencedModule, out var referencingModules))
            {
                return;
            }

            foreach (var module in referencingModules.Keys)
            {
                if (_referencesFrom.TryGetValue(module, out var referencedModules))
                {
                    referencedModules.TryRemove(referencedModule, out var _);
                }
            }
        }

        public override void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule)
        {
            if (!_referencesFrom.TryRemove(referencingModule, out var referencedModules))
            {
                return;
            }

            foreach (var module in referencedModules.Keys)
            {
                if (_referencesTo.TryGetValue(module, out var referencingModules))
                {
                    referencingModules.TryRemove(referencingModule, out var _);
                }
            }
        }

        public override IReadOnlyCollection<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule)
        {
            return _referencesFrom.TryGetValue(referencingModule, out var referencedModules) 
                        ? referencedModules.Keys.ToHashSet().AsReadOnly()
                        : new HashSet<QualifiedModuleName>().AsReadOnly();
        }

        public override IReadOnlyCollection<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule)
        {
            return _referencesTo.TryGetValue(referencedModule, out var referencingModules)
                        ? referencingModules.Keys.ToHashSet().AsReadOnly()
                        : new HashSet<QualifiedModuleName>().AsReadOnly();
        }

        
    }
}
