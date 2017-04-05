using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public abstract class ModuleToModuleReferenceManagerBase : IModuleToModuleReferenceManager
    {
        public abstract void AddModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule);

        public abstract void RemoveModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule);

        public virtual void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule)
        {
            var referencedModules = ModulesReferencedBy(referencingModule);
            foreach (var referencedModule in referencedModules)
            {
                RemoveModuleToModuleReference(referencedModule, referencingModule);
            }
        }

        public virtual void ClearModuleToModuleReferencesFromModule(IEnumerable<QualifiedModuleName> referencingModules)
        {
            foreach (var referencingModule in referencingModules)
            {
                ClearModuleToModuleReferencesFromModule(referencingModule);
            }
        }


        public virtual void ClearModuleToModuleReferencesToModule(QualifiedModuleName referencedModule)
        {
            var referencingModules = ModulesReferencing(referencedModule);
            foreach (var referencingModule in referencingModules)
            {
                RemoveModuleToModuleReference(referencedModule, referencingModule);
            }
        }

        public virtual void ClearModuleToModuleReferencesToModule(IEnumerable<QualifiedModuleName> referencedModules)
        {
            foreach (var referencedModule in referencedModules)
            {
                ClearModuleToModuleReferencesToModule(referencedModule);
            }
        }


        public abstract ICollection<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule);

        public virtual ICollection<QualifiedModuleName> ModulesReferencedByAny(IEnumerable<QualifiedModuleName> referencingModules)
        {
            var toModules = new HashSet<QualifiedModuleName>();

            foreach (var referencingModule in referencingModules)
            {
                toModules.UnionWith(ModulesReferencedBy(referencingModule));
            }
            return toModules;
        }


        public abstract ICollection<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule);

        public ICollection<QualifiedModuleName> ModulesReferencingAny(IEnumerable<QualifiedModuleName> referencedModules)
        {
            var fromModules = new HashSet<QualifiedModuleName>();

            foreach (var referencedModule in referencedModules)
            {
                fromModules.UnionWith(ModulesReferencing(referencedModule));
            }
            return fromModules;
        }

    }
}
