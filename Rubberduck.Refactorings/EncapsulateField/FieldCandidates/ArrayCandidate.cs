using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class ArrayCandidate : EncapsulateFieldCandidate
    {
        private string _subscripts;
        public ArrayCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            :base(declaration, validator)
        {
            ImplementLet = false;
            ImplementSet = false;
            AsTypeName_Field = declaration.AsTypeName;
            AsTypeName_Property = Tokens.Variant;
            CanBeReadWrite = false;
            IsReadOnly = true;

            _subscripts = string.Empty;
            if (declaration.Context.TryGetChildContext<VBAParser.SubscriptsContext>(out var ctxt))
            {
                _subscripts = ctxt.GetText();
            }
        }

        public override string AsUDTMemberDeclaration
            => $"{PropertyName}({_subscripts}) {Tokens.As} {AsTypeName_Field}";

        public override void LoadFieldReferenceContextReplacements()
        {
            foreach (var idRef in Declaration.References)
            {
                var replacementText = RequiresAccessQualification(idRef)
                    ? $"{QualifiedModuleName.ComponentName}.{ReferenceForPreExistingReferences}"
                    : FieldIdentifier;

                SetReferenceRewriteContent(idRef, replacementText);
            }
        }

        protected override void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
        {
            var context = idRef.Context;
            if (idRef.Context is VBAParser.IndexExprContext idxExpression)
            {
                context = idxExpression.children.ElementAt(0) as ParserRuleContext;
            }
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (context, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (context, replacementText));
        }
    }
}
