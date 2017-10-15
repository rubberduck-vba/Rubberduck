using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel
    {
        public RubberduckParserState State { get; }

        public Declaration TargetDeclaration { get; }

        public string PropertyName { get; set; }
        public string ParameterName { get; set; }
        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }
        public bool CanImplementLet { get; set; }

        public EncapsulateFieldModel(RubberduckParserState state, QualifiedSelection selection)
        {
            State = state;
            IList<Declaration> declarations = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).ToList();

            TargetDeclaration = declarations.FindVariable(selection);
        }
    }
}
