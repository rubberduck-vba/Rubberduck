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
        private readonly RubberduckParserState _state;
        public RubberduckParserState State { get { return _state; } }

        public Declaration TargetDeclaration { get; private set; }

        public string PropertyName { get; set; }
        public string ParameterName { get; set; }
        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }

        public EncapsulateFieldModel(RubberduckParserState state, QualifiedSelection selection)
        {
            _state = state;
            IList<Declaration> declarations = state.AllDeclarations
                                                        .Where(d => !d.IsBuiltIn && d.DeclarationType == DeclarationType.Variable)
                                                        .ToList();

            TargetDeclaration = declarations.FindVariable(selection);
        }
    }
}