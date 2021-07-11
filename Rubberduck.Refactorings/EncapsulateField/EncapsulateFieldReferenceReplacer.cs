using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldReferenceReplacer
    {
        void ReplaceReferences<T>(IEnumerable<T> selectedCandidates, 
            IRewriteSession rewriteSession) where T : IEncapsulateFieldCandidate;
    }
    /// <summary>
    /// EncapsulateFieldReferenceReplacer determines the replacement expressions for existing references
    /// of encapsulated fields.  It supports both direct encapsulation and encapsulation after wrapping the
    /// target field in a Private UserDefinedType
    /// </summary>
    public class EncapsulateFieldReferenceReplacer : IEncapsulateFieldReferenceReplacer
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IPropertyAttributeSetsGenerator _propertyAttributeSetsGenerator;
        private readonly IUDTMemberReferenceProvider _udtMemberReferenceProvider;
        private readonly Dictionary<IdentifierReference, (ParserRuleContext, string)> _identifierReplacements = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public EncapsulateFieldReferenceReplacer(IDeclarationFinderProvider declarationFinderProvider,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator,
            IUDTMemberReferenceProvider udtMemberReferenceProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _propertyAttributeSetsGenerator = propertyAttributeSetsGenerator;
            _udtMemberReferenceProvider = udtMemberReferenceProvider;
        }

        public void ReplaceReferences<T>(IEnumerable<T> selectedCandidates, IRewriteSession rewriteSession) where T : IEncapsulateFieldCandidate
        {
            if (!selectedCandidates.Any())
            {
                return;
            }

            ResolveReferenceContextReplacements(selectedCandidates);

            foreach (var kvPair in _identifierReplacements)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(kvPair.Key.QualifiedModuleName);
                (ParserRuleContext Context, string Text) = kvPair.Value;
                rewriter.Replace(Context, Text);
            }
        }

        private void ResolveReferenceContextReplacements<T>(IEnumerable<T> selectedCandidates ) where T : IEncapsulateFieldCandidate
        {
            var selectedVariableDeclarations = selectedCandidates.Select(sc => sc.Declaration).Cast<VariableDeclaration>();

            var udtFieldToMemberReferences = _udtMemberReferenceProvider.UdtFieldToMemberReferences(_declarationFinderProvider, selectedVariableDeclarations);

            foreach (var field in selectedCandidates)
            {
                if (IsInstanceOfPrivateUDT(field))
                {
                    if (udtFieldToMemberReferences.TryGetValue(field.Declaration, out var relevantReferences))
                    {
                        ResolvePrivateUDTMemberReferenceReplacements(field, relevantReferences);
                    }
                    continue;
                }
                
                ResolveNonPrivateUDTFieldReferenceReplacements(field);
            }
        }

        private void ResolvePrivateUDTMemberReferenceReplacements<T>(T field, IEnumerable<IdentifierReference> udtMemberReferencesToChange) where T: IEncapsulateFieldCandidate
        {
            if (field is IEncapsulateFieldAsUDTMemberCandidate wrappedField)
            {
                var wrappedField_WithStmtContexts = wrappedField.Declaration.References
                    .Where(rf => rf.Context.Parent.Parent is VBAParser.WithStmtContext)
                    .Select(rf => (rf, rf.Context));

                foreach ((IdentifierReference idRef, ParserRuleContext prCtxt) in wrappedField_WithStmtContexts)
                {
                    AddIdentifierReplacement(idRef, prCtxt, $"{wrappedField.ObjectStateUDT.FieldIdentifier}.{wrappedField.PropertyIdentifier}");
                }
            }

            foreach (var paSet in _propertyAttributeSetsGenerator.GeneratePropertyAttributeSets(field))
            {
                foreach (var rf in paSet.Declaration.References.Where(idRef => udtMemberReferencesToChange.Contains(idRef)))
                {
                    (ParserRuleContext context, string expression) = GenerateUDTMemberReplacementTuple(field, rf, paSet);
                    AddIdentifierReplacement(rf, context, expression);
                }
            }
        }

        private void ResolveNonPrivateUDTFieldReferenceReplacements<T>(T field) where T: IEncapsulateFieldCandidate
        {
            foreach (var idRef in field.Declaration.References.Where(rf => !(rf.IsArrayAccess || rf.IsDefaultMemberAccess)))
            {
                var replacementExpression = MustAccessUsingBackingField(idRef, field)
                    ? GetBackingIdentifier(field)
                    : field.PropertyIdentifier;

                if (RequiresModuleQualification(field, idRef, _declarationFinderProvider))
                {
                    replacementExpression = $"{field.QualifiedModuleName.ComponentName}.{replacementExpression}";
                }

                AddIdentifierReplacement(idRef, idRef.Context, replacementExpression);
            }
        }

        private static (ParserRuleContext, string) GenerateUDTMemberReplacementTuple<T>(T field, IdentifierReference rf, PropertyAttributeSet paSet) where T : IEncapsulateFieldCandidate
        {
            var replacementToken = paSet.PropertyName;
            if (rf.IsAssignment && field.IsReadOnly)
            {
                replacementToken = paSet.BackingField;
            }

            switch (rf.Context.Parent)
            {
                case VBAParser.WithMemberAccessExprContext wmaec:
                    return (wmaec, rf.IsAssignment && field.IsReadOnly ? $".{paSet.PropertyName}" : paSet.PropertyName);
                case VBAParser.MemberAccessExprContext maec:
                    return (maec, rf.IsAssignment && field.IsReadOnly ? replacementToken : paSet.PropertyName);
                default:
                    return (rf.Context, replacementToken);
            }
        }

        private static bool RequiresModuleQualification<T>(T field, IdentifierReference idRef, IDeclarationFinderProvider declarationFinderProvider) where T: IEncapsulateFieldCandidate
        {
            if (idRef.QualifiedModuleName == field.QualifiedModuleName)
            {
                return false;
            }

            var isUDTField = field.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false;

            return (isUDTField && !EncapsulateFieldUtilities.IsModuleQualifiedExternalReferenceOfUDTField(declarationFinderProvider, idRef, field.QualifiedModuleName))
                || !(idRef.Context.IsDescendentOf<VBAParser.MemberAccessExprContext>() || idRef.Context.IsDescendentOf<VBAParser.WithMemberAccessExprContext>());
        }

        private static bool IsInstanceOfPrivateUDT<T>(T field) where T: IEncapsulateFieldCandidate
        {
            bool IsPrivateUDT(IUserDefinedTypeCandidate u) => u.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private;
            
            return field is IEncapsulateFieldAsUDTMemberCandidate wrappedField
                ? wrappedField.WrappedCandidate is IUserDefinedTypeCandidate wrappedUDT && IsPrivateUDT(wrappedUDT)
                : field is IUserDefinedTypeCandidate udt && IsPrivateUDT(udt);
        }

        private static bool MustAccessUsingBackingField(IdentifierReference rf, IEncapsulateFieldCandidate field)
            => rf.QualifiedModuleName == field.QualifiedModuleName
                && ((rf.IsAssignment && field.IsReadOnly) || field.Declaration.IsArray);

        private static string GetBackingIdentifier(IEncapsulateFieldCandidate field)
        {
            var objStateUDT = field is IEncapsulateFieldAsUDTMemberCandidate udtM ? udtM.ObjectStateUDT : null;
            return objStateUDT is null
                ? field.BackingIdentifier
                : $"{objStateUDT.FieldIdentifier}.{field.BackingIdentifier}";
        }

        private void AddIdentifierReplacement(IdentifierReference idRef, ParserRuleContext context, string replacementText)
        {
            if (_identifierReplacements.ContainsKey(idRef))
            {
                _identifierReplacements[idRef] = (context, replacementText);
                return;
            }
            _identifierReplacements.Add(idRef, (context, replacementText));
        }
    }
}
