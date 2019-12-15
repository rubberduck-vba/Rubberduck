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
        bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT);
        bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT);
    }

    public class EncapsulateFieldNamesValidator : IEncapsulateFieldNamesValidator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private List<IEncapsulateFieldCandidate> FieldCandidates { set; get; }

        private List<IEncapsulateFieldCandidate> FlaggedCandidates => FieldCandidates.Where(rc => rc.EncapsulateFlag).ToList();

        public EncapsulateFieldNamesValidator(IDeclarationFinderProvider declarationFinderProvider, Func<List<IEncapsulateFieldCandidate>> retrieveCandidateFields = null)
        {
            _declarationFinderProvider = declarationFinderProvider;
            FieldCandidates = new List<IEncapsulateFieldCandidate>();
        }

        public void RegisterFieldCandidate(IEncapsulateFieldCandidate candidate)
        {
            FieldCandidates.Add(candidate);
        }

        private DeclarationFinder DeclarationFinder => _declarationFinderProvider.DeclarationFinder;

        public bool IsValidVBAIdentifier(string identifier, DeclarationType declarationType) 
            => VBAIdentifierValidator.IsValidIdentifier(identifier, declarationType);

        private List<Declaration> GetPotentialConflictMembers(IEncapsulateFieldCandidate candidate)
        {
            return _declarationFinderProvider.DeclarationFinder.Members(candidate.QualifiedModuleName)
                .Where(d => d != candidate.Declaration
                    && !IsAlwaysIgnoreNameConflictType(d)).ToList();
        }

        private List<Declaration> GetPotentialConflictMembers(IStateUDT stateUDT)
        {
            return _declarationFinderProvider.DeclarationFinder.Members(stateUDT.QualifiedModuleName)
                .Where(d => !IsAlwaysIgnoreNameConflictType(d)).ToList();
        }

        public bool HasConflictingPropertyIdentifier(IEncapsulateFieldCandidate candidate)
        {
            var members = GetPotentialConflictMembers(candidate);
            var newContentNames = FlaggedCandidates.Where(fc => fc.Declaration != candidate.Declaration).Select(fc => fc.PropertyName);
            newContentNames = newContentNames.Concat(FlaggedCandidates.Where(fc => fc.Declaration != candidate.Declaration).Select(fc => fc.FieldIdentifier));

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(candidate.PropertyName))
                || newContentNames.Any(ed => ed.EqualsVBAIdentifier(candidate.PropertyName));
        }

        public bool HasConflictingFieldIdentifier(IEncapsulateFieldCandidate candidate)
        {
            var members = GetPotentialConflictMembers(candidate);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(candidate.FieldIdentifier));
        }

        public bool IsConflictingFieldIdentifier(string fieldName, IEncapsulateFieldCandidate candidate)
        {
            var members = GetPotentialConflictMembers(candidate);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(fieldName) || candidate.PropertyName.EqualsVBAIdentifier(fieldName));
        }

        public bool IsConflictingStateUDTTypeIdentifier(IStateUDT stateUDT)
        {
            var members = GetPotentialConflictMembers(stateUDT);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(stateUDT.TypeIdentifier));
        }

        public bool IsConflictingStateUDTFieldIdentifier(IStateUDT stateUDT)
        {
            var members = GetPotentialConflictMembers(stateUDT);

            return members.Any(m => m.IdentifierName.EqualsVBAIdentifier(stateUDT.FieldIdentifier));
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

        private bool IsAlwaysIgnoreNameConflictType(Declaration d)
        {
            //TODO: // candidate.DeclarationType == DeclarationType.EnumerationMember || candidate.DeclarationType == DeclarationType.UserDefinedTypeMember;
            //FIXME: once a failing test is written
            return d.IsLocalVariable()
                    || d.IsLocalConstant()
                    || d.DeclarationType.HasFlag(DeclarationType.Parameter);
        }

        //private bool UsesScopeResolution(Antlr4.Runtime.RuleContext ruleContext)
        //{
        //    return (ruleContext is VBAParser.WithMemberAccessExprContext)
        //        || (ruleContext is VBAParser.MemberAccessExprContext);
        //}

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
