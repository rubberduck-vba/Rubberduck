using System.Collections.Generic;
using Rubberduck.VBEditor;
using System.Collections.Concurrent;

namespace Rubberduck.Parsing.VBA
{
    public class ModuleToModuleReferenceManager : ModuleToModuleReferenceManagerBase
    {
        private ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>> _referencesFrom = new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>>();
        private ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>> _referencesTo = new ConcurrentDictionary<QualifiedModuleName, ConcurrentDictionary<QualifiedModuleName, byte>>();


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
            byte dummyOutValue;
            ConcurrentDictionary<QualifiedModuleName, byte> referencedModules;
            if (_referencesFrom.TryGetValue(referencingModule, out referencedModules))
            {
                referencedModules.TryRemove(referencedModule, out dummyOutValue);
            }

            ConcurrentDictionary<QualifiedModuleName, byte> referencingModules;
            if (_referencesTo.TryGetValue(referencedModule, out referencingModules))
            {
                referencingModules.TryRemove(referencingModule, out dummyOutValue);
            }
        }

        public override void ClearModuleToModuleReferencesToModule(QualifiedModuleName referencedModule)
        {
            ConcurrentDictionary<QualifiedModuleName, byte> referencingModules;
            if (_referencesTo.TryRemove(referencedModule, out referencingModules))
            {
                byte dummyOutValue;
                ConcurrentDictionary<QualifiedModuleName, byte> referencedModules;
                foreach (var module in referencingModules.Keys)
                {
                    if(_referencesFrom.TryGetValue(module, out referencedModules))
                    {
                        referencedModules.TryRemove(referencedModule, out dummyOutValue);
                    }
                }
            }
        }

        public override void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule)
        {
            ConcurrentDictionary<QualifiedModuleName, byte> referencedModules;
            if (_referencesFrom.TryRemove(referencingModule, out referencedModules))
            {
                byte dummyOutValue;
                ConcurrentDictionary<QualifiedModuleName, byte> referencingModules;
                foreach (var module in referencedModules.Keys)
                {
                    if (_referencesTo.TryGetValue(module, out referencingModules))
                    {
                        referencingModules.TryRemove(referencingModule, out dummyOutValue);
                    }
                }
            }
        }

        public override IReadOnlyCollection<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule)
        {
            ConcurrentDictionary<QualifiedModuleName, byte> referencedModules;
            return _referencesFrom.TryGetValue(referencingModule, out referencedModules) 
                        ? referencedModules.Keys.ToHashSet().AsReadOnly()
                        : new HashSet<QualifiedModuleName>().AsReadOnly();
        }

        public override IReadOnlyCollection<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule)
        {
            ConcurrentDictionary<QualifiedModuleName, byte> referencingModules;
            return _referencesTo.TryGetValue(referencedModule, out referencingModules)
                        ? referencingModules.Keys.ToHashSet().AsReadOnly()
                        : new HashSet<QualifiedModuleName>().AsReadOnly();
        }

        
    }
}
