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
    //public interface IDeclarationFacade
    //{
    //    DeclarationType DeclarationType { get; }
    //    string IdentifierName { get;}
    //    string AsTypeName { get; }
    //    Accessibility Accessibility { get; }
    //    QualifiedModuleName QualifiedModuleName { get; }
    //    IEnumerable<IdentifierReference> References { get; }
    //}

    //public struct WrappedDeclaration : IDeclarationFacade
    //{
    //    public WrappedDeclaration(Declaration declaration)
    //    {
    //        DeclarationType = declaration.DeclarationType;
    //        IdentifierName = declaration.IdentifierName;
    //        Accessibility = declaration.Accessibility;
    //        References = declaration.References;
    //        AsTypeName = declaration.AsTypeName;
    //        QualifiedModuleName = declaration.QualifiedModuleName;
    //    }

    //    public DeclarationType DeclarationType { set; get; }
    //    public string IdentifierName { set; get; }
    //    public string AsTypeName { set; get; }
    //    public Accessibility Accessibility { set; get; }
    //    public QualifiedModuleName QualifiedModuleName { set; get; }
    //    public IEnumerable<IdentifierReference> References { set; get; }
    //}

    //public struct ProposedDeclaration : IDeclarationFacade
    //{

    //    public ProposedDeclaration(IEncapsulatedFieldDeclaration efd, DeclarationType declarationType)
    //    {
    //        DeclarationType = declarationType;
    //        IdentifierName = efd.PropertyName;
    //        Accessibility = Accessibility.Public;
    //        References = efd.References;
    //        if (declarationType.Equals(DeclarationType.Variable))
    //        {
    //            IdentifierName = efd.NewFieldName;
    //            Accessibility = Accessibility.Private;
    //            References = Enumerable.Empty<IdentifierReference>();
    //        }
    //        AsTypeName = efd.AsTypeName;
    //        QualifiedModuleName = efd.QualifiedModuleName;
    //    }

    //    public DeclarationType DeclarationType { set;  get; }
    //    public string IdentifierName { set;  get; }
    //    public string AsTypeName { set;  get; }
    //    public Accessibility Accessibility { set;  get; }
    //    public QualifiedModuleName QualifiedModuleName { set;  get; }
    //    public IEnumerable<IdentifierReference> References { set; get; }
    //}

    public interface IEncapsulateFieldNamesValidator
    {
        bool HasValidEncapsulationAttributes(IFieldEncapsulationAttributes attributes, QualifiedModuleName qmn, Predicate<Declaration> ignore);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private Func<IEnumerable<IEncapsulateFieldCandidate>> _selectedFieldsRetriever;
        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider, Func<IEnumerable<IEncapsulateFieldCandidate>> selectedFieldsRetriever = null)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedFieldsRetriever = selectedFieldsRetriever;
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder; 

        public bool HasValidEncapsulationAttributes(IFieldEncapsulationAttributes attributes, QualifiedModuleName qmn, Predicate<Declaration> ignore)
        {
            var isSelfConsistent = HasValidIdentifiers(attributes)
                && !HasInternalNameConflicts(attributes);

            if (!isSelfConsistent) { return false; }


            if (!attributes.FieldNameIsExemptFromValidation)
            {
                if (HasNewFieldNameConflicts(attributes, qmn, ignore) > 0) { return false; }
            }

            if (HasNewPropertyNameConflicts(attributes, qmn, ignore) > 0) { return false; }

            return true;
        }

        public int HasNewPropertyNameConflicts(IFieldEncapsulationAttributes attributes, QualifiedModuleName qmn, Predicate<Declaration> ignoreThisDeclaration)
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
                        && !ignoreThisDeclaration(match)
                        && !IsPrivateAccessiblityInOtherModule(match)
                        && !IsEnumOrUDTMemberDeclaration(match)
                        && !match.IsLocalVariable()).ToList();

            var candidates = new List<IEncapsulateFieldCandidate>();
            var candidateMatches = new List<IEncapsulateFieldCandidate>();
            var fields = _selectedFieldsRetriever is null ? Enumerable.Empty<IEncapsulateFieldCandidate>() : _selectedFieldsRetriever();
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
        public int HasNewFieldNameConflicts(IFieldEncapsulationAttributes attributes, QualifiedModuleName qmn, Predicate<Declaration> ignoreThisDeclaration)
        {
            var identifierMatches = DeclarationFinder.MatchName(attributes.NewFieldName)
                .Where(match => match.QualifiedModuleName == qmn
                        && !ignoreThisDeclaration(match)
                        && !IsEnumOrUDTMemberDeclaration(match)
                        && !match.IsLocalVariable()).ToList();

            var candidates = new List<IEncapsulateFieldCandidate>();
            var candidateMatches = new List<IEncapsulateFieldCandidate>();
            var fields = _selectedFieldsRetriever is null 
                ? Enumerable.Empty<IEncapsulateFieldCandidate>() 
                : _selectedFieldsRetriever();

            foreach (var efd in fields)
            {
                var matches = candidates.Where(c => c.NewFieldName.EqualsVBAIdentifier(efd.NewFieldName));
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

        private bool HasValidIdentifiers(IFieldEncapsulationAttributes attributes)
        {
            return VBAIdentifierValidator.IsValidIdentifier(attributes.NewFieldName, DeclarationType.Variable)
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
