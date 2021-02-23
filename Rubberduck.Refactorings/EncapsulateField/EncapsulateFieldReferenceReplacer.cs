using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldReferenceReplacer
    {
        void ReplaceReferences<T>(IEnumerable<T> selectedCandidates,
            IRewriteSession rewriteSession) where T : IEncapsulateFieldCandidate;
    }

    public class EncapsulateFieldReferenceReplacer : IEncapsulateFieldReferenceReplacer
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IPropertyAttributeSetsGenerator _propertyAttributeSetsGenerator;
        private readonly List<IEncapsulateFieldCandidate> _nonUDTCandidates = new List<IEncapsulateFieldCandidate>();
        private Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();
        private UserDefinedTypeInstanceProvider _udtInstanceProvider;

        public EncapsulateFieldReferenceReplacer(IDeclarationFinderProvider declarationFinderProvider,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _propertyAttributeSetsGenerator = propertyAttributeSetsGenerator;
        }

        public void ReplaceReferences<T>(IEnumerable<T> selectedCandidates, IRewriteSession rewriteSession) where T : IEncapsulateFieldCandidate
        {
            if (!selectedCandidates.Any())
            {
                return;
            }

            _udtInstanceProvider = new UserDefinedTypeInstanceProvider(_declarationFinderProvider, selectedCandidates.Select(c => c.Declaration));

            (List<T> selectedPrivateUDTFields, List<T> otherSelected) = SeparateFields(selectedCandidates);

            LoadReferenceReplacements(selectedPrivateUDTFields, otherSelected); 

            foreach (var replacement in IdentifierReplacements)
            {
                (ParserRuleContext Context, string Text) = replacement.Value;
                var rewriter = rewriteSession.CheckOutModuleRewriter(replacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }

        private static (List<T>, List<T>) SeparateFields<T>(IEnumerable<T> selectedCandidates) where T: IEncapsulateFieldCandidate
        {
            var pvtUDTs = Enumerable.Empty<T>();
            if (selectedCandidates.Any() && selectedCandidates.First() is IEncapsulateFieldAsUDTMemberCandidate)
            {
                pvtUDTs = selectedCandidates.Cast<IEncapsulateFieldAsUDTMemberCandidate>()
                    .Where(f => f.WrappedCandidate is IUserDefinedTypeCandidate udt
                        && udt.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private)
                    .Select(udtMemberCandidate => udtMemberCandidate)
                    .Cast<T>();

                return (pvtUDTs.ToList(), selectedCandidates.Except(pvtUDTs).ToList());
            }

            pvtUDTs = selectedCandidates
                .Where(f => f is IUserDefinedTypeCandidate udt
                    && udt.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private)
                .Select(udtMemberCandidate => udtMemberCandidate)
                .Cast<T>();
            return (pvtUDTs.ToList(), selectedCandidates.Except(pvtUDTs).ToList());
        }

        private void LoadReferenceReplacements<T>(List<T> selectedPrivateUDTFields, List<T> otherSelected) where T : IEncapsulateFieldCandidate
        {
            foreach (var field in selectedPrivateUDTFields)
            {
                foreach (var paSet in _propertyAttributeSetsGenerator.GeneratePropertyAttributeSets(field))
                {
                    AssignLocalUDTMemberReferenceExpressions(field, paSet);
                }
            }

            _nonUDTCandidates.AddRange(otherSelected.Cast<IEncapsulateFieldCandidate>());

            foreach (var field in otherSelected)
            {
                var isUDTField = field.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false;
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementExpression = MustAccessUsingBackingField(idRef, field)
                        ? GetBackingIdentifier(field)
                        : field.PropertyIdentifier;

                    if (idRef.QualifiedModuleName != field.QualifiedModuleName)
                    {
                        if (isUDTField)
                        {
                            if (idRef.Context.Parent is ParserRuleContext prCtxt && !IsModuleQualified(prCtxt, field.QualifiedModuleName))
                            {
                                replacementExpression = $"{field.QualifiedModuleName.ComponentName}.{replacementExpression}";
                            }
                        }
                        else if (!(idRef.Context.IsDescendentOf<VBAParser.MemberAccessExprContext>() || idRef.Context.IsDescendentOf<VBAParser.WithMemberAccessExprContext>()))
                        {
                            replacementExpression = $"{field.QualifiedModuleName.ComponentName}.{replacementExpression}";
                        }
                    }

                    if (!(idRef.IsArrayAccess || idRef.IsDefaultMemberAccess))
                    {
                        AddIdentifierReplacement(idRef, idRef.Context, replacementExpression);
                    }
                }
            }
        }

        //TODO: Write tests using Member and With access contexts locally and externally - replacement text is different for readOnly scenario
        private void AssignLocalUDTMemberReferenceExpressions<T>(T field, PropertyAttributeSet paSet) where T: IEncapsulateFieldCandidate
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

            var nonPropertyAssignmentsLHS = field.IsReadOnly ? paSet.BackingField : paSet.PropertyName;
            var udtInstance = _udtInstanceProvider[field.Declaration as VariableDeclaration];

            foreach (var rf in paSet.Declaration.References)
            {
                if (!udtInstance.UDTMemberReferences.Contains(rf))
                {
                    continue;
                }


                var expression = rf.IsAssignment ? nonPropertyAssignmentsLHS : paSet.PropertyName;
                if (rf.Context.Parent is VBAParser.WithMemberAccessExprContext wmaec)
                {
                    expression = field.IsReadOnly && rf.IsAssignment ? $".{paSet.PropertyName}" : paSet.PropertyName;
                    AddIdentifierReplacement(rf, wmaec, expression);
                    continue;
                }

                if (rf.Context.Parent is VBAParser.MemberAccessExprContext maec)
                {
                    if (rf.IsAssignment && field.IsReadOnly)
                    {
                        AddIdentifierReplacement(rf, maec, expression);
                    }
                    else
                    {
                        AddIdentifierReplacement(rf, maec, paSet.PropertyName);
                    }
                    continue;
                }

                AddIdentifierReplacement(rf, rf.Context, expression);
            }
        }

        //TODO: this predicate function needs be validated
        private static bool IsModuleQualified(ParserRuleContext ctxt, QualifiedModuleName qmn)
        {
            return qmn.ComponentType == ComponentType.ClassModule || ctxt.GetText().StartsWith(qmn.ComponentName) || ctxt.Parent.GetText().StartsWith(qmn.ComponentName)
                || ctxt is VBAParser.WithMemberAccessExprContext;
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
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (context, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (context, replacementText));
        }
    }
}
