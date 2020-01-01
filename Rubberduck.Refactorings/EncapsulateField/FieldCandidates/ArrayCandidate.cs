using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IArrayCandidate : IEncapsulateFieldCandidate
    {
    }

    public class ArrayCandidate : EncapsulateFieldCandidate, IArrayCandidate
    {
        private string _subscripts;
        public ArrayCandidate(Declaration declaration, IValidateEncapsulateFieldNames validator)
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

        private bool HasExternalRedimOperation
            => Declaration.References.Any(rf => rf.QualifiedModuleName != QualifiedModuleName
                    && rf.Context.TryGetAncestor<VBAParser.RedimVariableDeclarationContext>(out _));


        public override string AsUDTMemberDeclaration
            => $"{PropertyName}({_subscripts}) {Tokens.As} {AsTypeName_Field}";

        public override bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!EncapsulateFlag) { return true; }

            if (HasExternalRedimOperation)
            {
                errorMessage = string.Format(EncapsulateFieldResources.ArrayHasExternalRedimFormat, Declaration.IdentifierName, Declaration.QualifiedModuleName.ComponentName);
                return false;
            }

            if (ConvertFieldToUDTMember)
            {
                return TryValidateAsUDTMemberEncapsulationAttributes(out errorMessage, true);
            }

            if (!TryValidateEncapsulationAttributes(DeclarationType.Property, out errorMessage, true))
            {
                return false;
            }

            if (NamesValidator.HasConflictingIdentifier(this, DeclarationType.Variable, out errorMessage))
            {
                return false;
            }
            return true;
        }

        public override void LoadFieldReferenceContextReplacements(string referenceQualifier = null)
        {
            ReferenceQualifier = referenceQualifier;
            foreach (var idRef in Declaration.References)
            {
                //Locally, we do all operations using the backing field
                if (idRef.QualifiedModuleName == QualifiedModuleName)
                {
                    var accessor = ConvertFieldToUDTMember 
                        ? ReferenceForPreExistingReferences 
                        : FieldIdentifier;

                    SetReferenceRewriteContent(idRef, accessor);
                    continue;
                }

                var replacementText = RequiresAccessQualification(idRef)
                    ? $"{QualifiedModuleName.ComponentName}.{ReferenceForPreExistingReferences}"
                    : ReferenceForPreExistingReferences;

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
