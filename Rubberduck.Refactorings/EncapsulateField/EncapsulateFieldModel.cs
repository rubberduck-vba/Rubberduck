using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel
    {
        public RubberduckParserState State { get; }

        public EncapsulateFieldModel(RubberduckParserState state, QualifiedSelection selection)
        {
            State = state;
            IList<Declaration> declarations = state.DeclarationFinder.UserDeclarations(DeclarationType.Variable).ToList();

            TargetDeclaration = declarations.FindVariable(selection);
            ParameterName = "value";
        }

        private Declaration _targetDeclaration;
        public Declaration TargetDeclaration
        {
            get => _targetDeclaration;
            set {
                _targetDeclaration = value;
                AssignSetterAndLetterAvailability();
            }
        }

        public string PropertyName { get; set; }
        public string ParameterName { get; set; }
        public bool ImplementLetSetterType { get; set; }
        public bool ImplementSetSetterType { get; set; }
        public bool CanImplementLet { get; private set; }
        public bool CanImplementSet { get; private set; }

        private void AssignSetterAndLetterAvailability()
        {
            var isVariant = _targetDeclaration.AsTypeName.Equals(Tokens.Variant);
            var isValueType = !isVariant && (SymbolList.ValueTypes.Contains(_targetDeclaration.AsTypeName) ||
                                             _targetDeclaration.DeclarationType == DeclarationType.Enumeration);

            if (_targetDeclaration.References.Any(r => r.IsAssignment))
            {
                if (isVariant)
                {
                    RuleContext node = _targetDeclaration.References.First(r => r.IsAssignment).Context;
                    while (!(node is VBAParser.LetStmtContext) && !(node is VBAParser.SetStmtContext))
                    {
                        node = node.Parent;
                    }

                    if (node is VBAParser.LetStmtContext)
                    {
                        CanImplementLet = true;
                    }
                    else
                    {
                        CanImplementSet = true;
                    }
                }
                else if (isValueType)
                {
                    CanImplementLet = true;
                }
                else
                {
                    CanImplementSet = true;
                }
            }
            else
            {
                if (isValueType)
                {
                    CanImplementLet = true;
                }
                else if (!isVariant)
                {
                    CanImplementSet = true;
                }
                else
                {
                    CanImplementLet = true;
                    CanImplementSet = true;
                }
            }
        }
    }
}
