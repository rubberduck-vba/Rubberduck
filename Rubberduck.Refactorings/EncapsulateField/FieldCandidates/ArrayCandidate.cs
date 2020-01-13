using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IArrayCandidate : IEncapsulateFieldCandidate
    {
        string UDTMemberDeclaration { get;}
    }

    public class ArrayCandidate : EncapsulateFieldCandidate, IArrayCandidate
    {
        private string _subscripts;
        public ArrayCandidate(Declaration declaration, IValidateVBAIdentifiers validator)
            :base(declaration, validator)
        {
            ImplementLet = false;
            ImplementSet = false;
            PropertyAsTypeName = Tokens.Variant;
            CanBeReadWrite = false;
            IsReadOnly = true;

            _subscripts = string.Empty;
            if (declaration.Context.TryGetChildContext<VBAParser.SubscriptsContext>(out var ctxt))
            {
                _subscripts = ctxt.GetText();
            }
        }

        public override bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!EncapsulateFlag) { return true; }

            if (HasExternalRedimOperation)
            {
                errorMessage = string.Format(RubberduckUI.EncapsulateField_ArrayHasExternalRedimFormat, IdentifierName);
                return false;
            }
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
        }

        public string UDTMemberDeclaration
            => $"{PropertyIdentifier}({_subscripts}) {Tokens.As} {Declaration.AsTypeName}";

        protected override string AccessorInProperty
            => $"{BackingIdentifier}";

        protected override string AccessorLocalReference(IdentifierReference idRef)
            => $"{BackingIdentifier}";

        private bool HasExternalRedimOperation
            => Declaration.References.Any(rf => rf.QualifiedModuleName != QualifiedModuleName
                    && rf.Context.TryGetAncestor<VBAParser.RedimVariableDeclarationContext>(out _));
    }
}
