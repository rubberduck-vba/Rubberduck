using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.ReplaceReferences;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldReferenceReplacer
    {
        void ReplaceReferences<T>(IEnumerable<T> selectedCandidates,
            IRewriteSession rewriteSession,
            IObjectStateUDT objStateUDT = null) where T : IEncapsulateFieldCandidate; 
    }

    public class EncapsulateFieldReferenceReplacer : IEncapsulateFieldReferenceReplacer
    {
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replacePrivateUDTMemberReferencesModelFactory;
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replacePrivateUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceReferencesRefactoringAction;
        private readonly IPropertyAttributeSetsGenerator _propertyAttributeSetsGenerator;
        public EncapsulateFieldReferenceReplacer(IReplacePrivateUDTMemberReferencesModelFactory replacePrivateUDTMemberReferencesModelFactory,
            ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> replacePrivateUDTMemberReferencesRefactoringAction,
            ICodeOnlyRefactoringAction<ReplaceReferencesModel> replaceReferencesRefactoringAction,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator)
        {
            _replacePrivateUDTMemberReferencesModelFactory = replacePrivateUDTMemberReferencesModelFactory;
            _replacePrivateUDTMemberReferencesRefactoringAction = replacePrivateUDTMemberReferencesRefactoringAction;
            _replaceReferencesRefactoringAction = replaceReferencesRefactoringAction;
            _propertyAttributeSetsGenerator = propertyAttributeSetsGenerator;
        }

        public void ReplaceReferences<T>(IEnumerable<T> selectedCandidates, 
            IRewriteSession rewriteSession,
            IObjectStateUDT objStateUDT = null) where T : IEncapsulateFieldCandidate
        {
            (List<T> selectedPrivateUDTFields, List<T> otherSelected) = SeparateFields(selectedCandidates);
            
            if (selectedPrivateUDTFields.Any())
            {
                var replacePrivateUDTMemberReferencesModel = CreatePrivateUDTMemberReferencesModel(_replacePrivateUDTMemberReferencesModelFactory, selectedPrivateUDTFields, _propertyAttributeSetsGenerator);
                _replacePrivateUDTMemberReferencesRefactoringAction.Refactor(replacePrivateUDTMemberReferencesModel, rewriteSession);
            }

            if (otherSelected.Any())
            {
                var replaceReferencesModel = CreateReplaceReferencesModel(otherSelected, objStateUDT);
                _replaceReferencesRefactoringAction.Refactor(replaceReferencesModel, rewriteSession);
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
        private static ReplacePrivateUDTMemberReferencesModel CreatePrivateUDTMemberReferencesModel<T>(IReplacePrivateUDTMemberReferencesModelFactory factory, IEnumerable<T> privateUDTFields, IPropertyAttributeSetsGenerator paSetGenerator) where T : IEncapsulateFieldCandidate
        {
            var model = factory.Create(privateUDTFields.Select(f => f.Declaration).Cast<VariableDeclaration>());

            foreach (var field in privateUDTFields)
            {
                InitializeReplacePrivateUDTMemberReferencesModel(paSetGenerator.GeneratePropertyAttributeSets(field), model, field);
            }
            return model;
        }

        private static ReplaceReferencesModel CreateReplaceReferencesModel<T>(IEnumerable<T> nonPrivateUDTFields, IObjectStateUDT objectStateUDTField = null) where T : IEncapsulateFieldCandidate
        {
            var replaceReferencesModel = new ReplaceReferencesModel()
            {
                ModuleQualifyExternalReferences = true,
            };

            InitializeReplaceReferencesModel(replaceReferencesModel, nonPrivateUDTFields, objectStateUDTField);
            return replaceReferencesModel;
        }
        private static void InitializeReplaceReferencesModel<T>(ReplaceReferencesModel model, IEnumerable<T> fields, IObjectStateUDT objStateUDT = null) where T : IEncapsulateFieldCandidate
        {
            foreach (var field in fields)
            {
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementExpression = MustAccessUsingBackingField(idRef, field)
                        ? GetBackingIdentifier(field, objStateUDT)
                        : field.PropertyIdentifier;

                    model.RegisterReferenceReplacementExpression(idRef, replacementExpression);
                }
            }
        }
        private static void InitializeReplacePrivateUDTMemberReferencesModel<T>(IEnumerable<PropertyAttributeSet> propertyAttributeSets, ReplacePrivateUDTMemberReferencesModel model, T candidate) where T : IEncapsulateFieldCandidate
        {
            foreach (var paSet in propertyAttributeSets)
            {
                var nonPropertyAssignmentsLHS = candidate.IsReadOnly ? paSet.BackingField : paSet.PropertyName;
                foreach (var rf in paSet.Declaration.References)
                {
                    var expression = rf.IsAssignment ? nonPropertyAssignmentsLHS : paSet.PropertyName;
                    model.RegisterReferenceReplacementExpression(rf, expression);
                }
            }
        }
        private static bool MustAccessUsingBackingField(IdentifierReference rf, IEncapsulateFieldCandidate field)
            => rf.QualifiedModuleName == field.QualifiedModuleName
                && ((rf.IsAssignment && field.IsReadOnly) || field.Declaration.IsArray);

        private static string GetBackingIdentifier(IEncapsulateFieldCandidate field, IObjectStateUDT objStateUDT = null)
        {
            return objStateUDT is null
                ? field.BackingIdentifier
                : $"{objStateUDT.FieldIdentifier}.{field.BackingIdentifier}";
        }
    }
}
