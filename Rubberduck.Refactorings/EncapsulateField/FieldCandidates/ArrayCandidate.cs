using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
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
            FieldAsTypeName = declaration.AsTypeName;
            PropertyAsTypeName = Tokens.Variant;
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

        public override bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
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

        public override string AccessorInProperty
            => $"{ObjectStateUDT.FieldIdentifier}.{UDTMemberIdentifier}";

        public override string AccessorLocalReference
            => $"{ObjectStateUDT.FieldIdentifier}.{UDTMemberIdentifier}";

        public override string UDTMemberDeclaration
            => $"{PropertyIdentifier}({_subscripts}) {Tokens.As} {Declaration.AsTypeName}";
    }
}
