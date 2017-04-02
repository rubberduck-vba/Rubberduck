using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public class StateModuleToModuleReferenceManager: ModuleToModuleReferenceManagerBase 
    {
        private readonly RubberduckParserState _state;

        public StateModuleToModuleReferenceManager(RubberduckParserState state)
        {
            _state = state;
        }


        public override void AddModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule)
        {
            _state.AddModuleToModuleReference(referencedModule, referencingModule);
        }

        public override void RemoveModuleToModuleReference(QualifiedModuleName referencedModule, QualifiedModuleName referencingModule)
        {
            _state.RemoveModuleToModuleReference(referencedModule, referencingModule);
        }

        public override void ClearModuleToModuleReferencesFromModule(QualifiedModuleName referencingModule)
        {
            _state.ClearModuleToModuleReferencesFromModule(referencingModule);
        }

        public override void ClearModuleToModuleReferencesToModule(QualifiedModuleName referencedModule)
        {
            _state.ClearModuleToModuleReferencesToModule(referencedModule);
        }

        public override ICollection<QualifiedModuleName> ModulesReferencedBy(QualifiedModuleName referencingModule)
        {
            return _state.ModulesReferencedBy(referencingModule);
        }

        public override ICollection<QualifiedModuleName> ModulesReferencing(QualifiedModuleName referencedModule)
        {
            return _state.ModulesReferencing(referencedModule);
        }
    }
}
