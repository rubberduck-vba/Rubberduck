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
    public interface IEncapsulateFieldValidator
    {
        void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate);
        bool HasValidEncapsulationAttributes(IEncapsulateFieldCandidate candidate, IEnumerable<Declaration> ignore);
        bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType);
        bool HasConflictingPropertyIdentifier(IEncapsulateFieldCandidate candidate);
        bool HasConflictingFieldIdentifier(IEncapsulateFieldCandidate candidate);
        bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate);
    }

    public interface IEncapsulateFieldNamesValidator : IEncapsulateFieldValidator
    {
        //bool HasIdentifierConflicts(string identifier, DeclarationType declarationType);
        bool IsConflictingStateUDTIdentifier(IUserDefinedTypeCandidate candidate);
        bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT);
        bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private List<IEncapsulateFieldCandidate> _registeredCandidates;

        private List<IEncapsulateFieldCandidate> FieldCandidates => _registeredCandidates; // _fieldCandidates.Value;

        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider, Func<List<IEncapsulateFieldCandidate>> retrieveCandidateFields = null)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _registeredCandidates = new List<IEncapsulateFieldCandidate>();
        }

        public void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate)
        {
            _registeredCandidates.Add(candidate);
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;

        public bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType) 
            => VBAIdentifierValidator.IsValidIdentifier(identifier, declarationType);

        public bool HasConflictingPropertyIdentifier(IEncapsulateFieldCandidate candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            var edits = FieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration != candidate.Declaration).Select(fc => fc.PropertyName);
            edits = edits.Concat(FieldCandidates.Where(fc => fc.EncapsulateFlag && fc.Declaration != candidate.Declaration).Select(fc => fc.FieldIdentifier));


            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(candidate.PropertyName))
                || edits.Any(ed => ed.EqualsVBAIdentifier(candidate.PropertyName));
        }

        public bool HasConflictingFieldIdentifier(IEncapsulateFieldCandidate candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(candidate.FieldIdentifier));
        }

        public bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(fieldName) || candidate.PropertyName.EqualsVBAIdentifier(fieldName));
        }

        public bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(stateUDT.TypeIdentifier));
        }

        public bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(stateUDT.FieldIdentifier));
        }


        public bool IsConflictingStateUDTIdentifier(IUserDefinedTypeCandidate candidate)
        {
            var members = _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(candidate.AsTypeName));
        }

        public bool HasValidEncapsulationAttributes(IEncapsulateFieldCandidate candidate, IEnumerable<Declaration> ignore)
        {
            if (!candidate.EncapsulateFlag) { return true; }

            var internalValidations = candidate as IEncapsulateFieldCandidateValidations;

            if (!internalValidations.IsSelfConsistent
                || HasConflictingPropertyIdentifier(candidate))
            {
                return false;
            }
            return true;
        }

        //public int HasNewPropertyNameConflicts(IEncapsulateFieldCandidate attributes, QualifiedModuleName qmn, IEnumerable<Declaration> declarationsToIgnore)
        //{
        //    Predicate<Declaration> IsPrivateAccessiblityInOtherModule = (Declaration dec) => dec.QualifiedModuleName != qmn && dec.Accessibility.Equals(Accessibility.Private);
        //    Predicate<Declaration> IsInSearchScope = null;
        //    if (qmn.ComponentType == ComponentType.ClassModule)
        //    {
        //        IsInSearchScope = (Declaration dec) => dec.QualifiedModuleName == qmn;
        //    }
        //    else
        //    {
        //        IsInSearchScope = (Declaration dec) => dec.QualifiedModuleName.ProjectId == qmn.ProjectId;
        //    }

        //    var identifierMatches = DeclarationFinder.MatchName(attributes.PropertyName)
        //        .Where(match => IsInSearchScope(match)
        //                && !declarationsToIgnore.Contains(match)
        //                && !IsPrivateAccessiblityInOtherModule(match)
        //                && !IsEnumOrUDTMemberDeclaration(match)
        //                && !match.IsLocalVariable()).ToList();

        //    var candidates = new List<IEncapsulateFieldCandidate>();
        //    var candidateMatches = new List<IEncapsulateFieldCandidate>();
        //    foreach (var efd in FieldCandidates)
        //    {
        //        var matches = candidates.Where(c => c.PropertyName.EqualsVBAIdentifier(efd.PropertyName));
        //        if (matches.Any())
        //        {
        //            candidateMatches.Add(efd);
        //        }
        //        candidates.Add(efd);
        //    }

        //    return identifierMatches.Count() + candidateMatches.Count();
        //}

        //public bool HasIdentifierConflicts(string identifier, DeclarationType declarationType)
        //{
        //    return true;
        //}


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

        private bool HasValidIdentifiers(IEncapsulateFieldCandidate attributes, DeclarationType declarationType)
        {
            return VBAIdentifierValidator.IsValidIdentifier(attributes.FieldIdentifier, declarationType)
                && VBAIdentifierValidator.IsValidIdentifier(attributes.PropertyName, DeclarationType.Property)
                && VBAIdentifierValidator.IsValidIdentifier(attributes.ParameterName, DeclarationType.Parameter);
        }

        private bool HasInternalNameConflicts(IEncapsulateFieldCandidate attributes)
        {
            return attributes.PropertyName.EqualsVBAIdentifier(attributes.FieldIdentifier)
                || attributes.PropertyName.EqualsVBAIdentifier(attributes.ParameterName)
                || attributes.FieldIdentifier.EqualsVBAIdentifier(attributes.ParameterName);
        }
    }
}
