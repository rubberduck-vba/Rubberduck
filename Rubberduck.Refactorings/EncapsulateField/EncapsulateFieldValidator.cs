using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldNamesValidator
    {
        bool HasValidEncapsulationAttributes(IEncapsulateFieldCandidate attributes, QualifiedModuleName qmn, IEnumerable<Declaration> ignore, DeclarationType declarationType);
        void ForceNonConflictEncapsulationAttributes(IEncapsulateFieldCandidate attributes, QualifiedModuleName qmn, Declaration target);
        void ForceNonConflictPropertyName(IEncapsulateFieldCandidate attributes, QualifiedModuleName qmn, Declaration target);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private Func<IEnumerable<IEncapsulateFieldCandidate>> _candidateFieldsRetriever;
        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider, Func<IEnumerable<IEncapsulateFieldCandidate>> selectedFieldsRetriever = null)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _candidateFieldsRetriever = selectedFieldsRetriever;
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;

        public void ForceNonConflictEncapsulationAttributes(IEncapsulateFieldCandidate candidate, QualifiedModuleName qmn, Declaration target)
        {
            if (target?.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember) ?? false)
            {
                ForceNonConflictPropertyName(candidate, qmn, target);
                return;
            }
            ForceNonConflictNewName(candidate, qmn, target);
        }

        public void ForceNonConflictNewName(IEncapsulateFieldCandidate candidate, QualifiedModuleName qmn, Declaration target)
        {
            var attributes = candidate.EncapsulationAttributes;
            var identifier = candidate.NewFieldName;
            var ignore = target is null ? Enumerable.Empty<Declaration>() : new Declaration[] { target };

            var isValidAttributeSet = HasValidEncapsulationAttributes(candidate, qmn, ignore);
            for (var idx = 1; idx < 9 && !isValidAttributeSet; idx++)
            {
                attributes.NewFieldName = $"{identifier}{idx}";
                isValidAttributeSet = HasValidEncapsulationAttributes(candidate, qmn, ignore);
            }
        }

        public void ForceNonConflictPropertyName(IEncapsulateFieldCandidate candidate, QualifiedModuleName qmn, Declaration target)
        {
            var attributes = candidate.EncapsulationAttributes;
            var identifier = attributes.PropertyName;
            var ignore = target is null ? Enumerable.Empty<Declaration>() : new Declaration[] { target };
            var isValidAttributeSet = HasValidEncapsulationAttributes(candidate, qmn, ignore);
            for (var idx = 1; idx < 9 && !isValidAttributeSet; idx++)
            {
                attributes.PropertyName = $"{identifier}{idx}";
                isValidAttributeSet = HasValidEncapsulationAttributes(candidate, qmn, ignore);
            }
        }

        public bool HasValidEncapsulationAttributes(IEncapsulateFieldCandidate candidate, QualifiedModuleName qmn, IEnumerable<Declaration> ignore, DeclarationType declaration = DeclarationType.Variable)
        {
            var attributes = candidate.EncapsulationAttributes;
            var hasValidIdentifiers = HasValidIdentifiers(attributes, declaration);
            var hasInternalNameConflicts = HasInternalNameConflicts(attributes);

            var isSelfConsistent = hasValidIdentifiers || hasInternalNameConflicts;

            if (!isSelfConsistent) { return false; }


            if (!candidate.FieldNameIsExemptFromValidation)
            {
                if (HasNewFieldNameConflicts(attributes, qmn, ignore) > 0) { return false; }
            }

            if (HasNewPropertyNameConflicts(attributes, qmn, ignore) > 0) { return false; }

            return true;
        }

        public int HasNewPropertyNameConflicts(IFieldEncapsulationAttributes attributes, QualifiedModuleName qmn, IEnumerable<Declaration> declarationsToIgnore)
        {
            Predicate<Declaration> IsPrivateAccessiblityInOtherModule = (Declaration dec) => dec.QualifiedModuleName != qmn && dec.Accessibility.Equals(Accessibility.Private);
            Predicate<Declaration> IsInSearchScope = null;
            if (qmn.ComponentType == ComponentType.ClassModule)
            {
                IsInSearchScope = (Declaration dec) => dec.QualifiedModuleName == qmn;
            }
            else
            {
                IsInSearchScope = (Declaration dec) => dec.QualifiedModuleName.ProjectId == qmn.ProjectId;
            }

            var identifierMatches = DeclarationFinder.MatchName(attributes.PropertyName)
                .Where(match => IsInSearchScope(match)
                        && !declarationsToIgnore.Contains(match)
                        && !IsPrivateAccessiblityInOtherModule(match)
                        && !IsEnumOrUDTMemberDeclaration(match)
                        && !match.IsLocalVariable()).ToList();

            var candidates = new List<IEncapsulateFieldCandidate>();
            var candidateMatches = new List<IEncapsulateFieldCandidate>();
            var fields = _candidateFieldsRetriever is null ? Enumerable.Empty<IEncapsulateFieldCandidate>() : _candidateFieldsRetriever();
            foreach (var efd in fields)
            {
                var matches = candidates.Where(c => c.PropertyName.EqualsVBAIdentifier(efd.PropertyName));
                if (matches.Any())
                {
                    candidateMatches.Add(efd);
                }
                candidates.Add(efd);
            }

            return identifierMatches.Count() + candidateMatches.Count();
        }

        //FieldNames are always Private, so only look within the same module as the field to encapsulate
        public int HasNewFieldNameConflicts(IFieldEncapsulationAttributes attributes, QualifiedModuleName qmn, IEnumerable<Declaration> declarationsToIgnore)
        {
            var rawmatches = DeclarationFinder.MatchName(attributes.NewFieldName);
            var identifierMatches = DeclarationFinder.MatchName(attributes.NewFieldName)
                .Where(match => match.QualifiedModuleName == qmn
                        && !declarationsToIgnore.Contains(match)
                        && !IsEnumOrUDTMemberDeclaration(match)
                        && !match.IsLocalVariable()).ToList();

            var candidates = new List<IEncapsulateFieldCandidate>();
            var candidateMatches = new List<IEncapsulateFieldCandidate>();
            var fields = _candidateFieldsRetriever is null 
                ? Enumerable.Empty<IEncapsulateFieldCandidate>() 
                : _candidateFieldsRetriever();

            foreach (var efd in fields)
            {
                var matches = candidates.Where(c => c.EncapsulateFlag &&  c.NewFieldName.EqualsVBAIdentifier(efd.NewFieldName)
                                                        || c.IdentifierName.EqualsVBAIdentifier(efd.NewFieldName));
                if (matches.Where(m => m.TargetID != efd.TargetID).Any())
                {
                    candidateMatches.Add(efd);
                }
                candidates.Add(efd);
            }

            return identifierMatches.Count() + candidateMatches.Count();
        }

        private bool IsEnumOrUDTMemberDeclaration(Declaration candidate)
        {
            return candidate.DeclarationType == DeclarationType.EnumerationMember
                       || candidate.DeclarationType == DeclarationType.UserDefinedTypeMember;
        }

        private bool UsesScopeResolution(Antlr4.Runtime.RuleContext ruleContext)
        {
            return (ruleContext is VBAParser.WithMemberAccessExprContext)
                || (ruleContext is VBAParser.MemberAccessExprContext);
        }

        private bool HasValidIdentifiers(IFieldEncapsulationAttributes attributes, DeclarationType declarationType)
        {
            return VBAIdentifierValidator.IsValidIdentifier(attributes.NewFieldName, declarationType)
                && VBAIdentifierValidator.IsValidIdentifier(attributes.PropertyName, DeclarationType.Property)
                && VBAIdentifierValidator.IsValidIdentifier(attributes.ParameterName, DeclarationType.Parameter);
        }

        private bool HasInternalNameConflicts(IFieldEncapsulationAttributes attributes)
        {
            return attributes.PropertyName.EqualsVBAIdentifier(attributes.NewFieldName)
                || attributes.PropertyName.EqualsVBAIdentifier(attributes.ParameterName)
                || attributes.NewFieldName.EqualsVBAIdentifier(attributes.ParameterName);
        }
    }
}
