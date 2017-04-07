using Rubberduck.VBEditor;
using System.Collections.Generic;


namespace Rubberduck.Parsing.VBA
{
    public interface IModuleToModuleReferenceManager
    {
        void AddModuleToModuleReference(QualifiedModuleName referencingModule, QualifiedModuleName referencedModule);
        void RemoveModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule);
        void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule);
        void ClearModuleToModuleReferencesFromModule(IEnumerable<QualifiedModuleName> referencingModules);
        void ClearModuleToModuleReferencesToModule(QualifiedModuleName referencedModule);
        void ClearModuleToModuleReferencesToModule(IEnumerable<QualifiedModuleName> referencedModules);

        ICollection<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule);
        ICollection<QualifiedModuleName> ModulesReferencedByAny(IEnumerable<QualifiedModuleName> referencingModules);
        ICollection<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule);
        ICollection<QualifiedModuleName> ModulesReferencingAny(IEnumerable<QualifiedModuleName> referencedModules);
    }
}
