using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Supporting base class for Subroutines, Functions, and Properties
    /// </summary>
    /// <seealso cref="ConflictFinderMembers"/>
    /// <seealso cref="ConflictFinderProperties"/>
    public abstract class ConflictFinderModuleCodeSection : ConflictFinderBase
    {
        public ConflictFinderModuleCodeSection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public abstract override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);

        protected bool TryFindMemberConflictChecksCommon(IConflictDetectionDeclarationProxy memberProxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesInTargetModule, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            if (!identifierMatchesInTargetModule.Any()) { return false; }

            if (ModuleLevelElementChecks(identifierMatchesInTargetModule, out var moduleLevelConflicts))
            {
                conflicts = AddConflicts(conflicts, memberProxy, moduleLevelConflicts);
            }

            if (LocalDeclarationsHaveSameNameAsParentScope(memberProxy, identifierMatchesInTargetModule, out var localConflicts))
            {
                conflicts = AddConflicts(conflicts, memberProxy, localConflicts);
            }

            if (memberProxy.DeclarationType.HasFlag(DeclarationType.Function)
                    && TryFindFunctionNameMatchesParameterConflict(memberProxy, identifierMatchesInTargetModule, out var parameterConflicts))
            {
                conflicts = AddConflicts(conflicts, memberProxy, parameterConflicts);
            }
            return conflicts.Values.Any();
        }

        protected bool NonModuleQualifiedMemberReferenceConflicts(IConflictDetectionDeclarationProxy memberProxy, IConflictDetectionSessionData sessionData, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesAllModules, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            if (!IsExistingTargetModule(memberProxy, out var targetModule))
            {
                return false;
            }

            var identifierMatchesExternal = identifierMatchesAllModules.Where(d => d.QualifiedModuleName != targetModule.QualifiedModuleName);

            var nonQualifiedMemberReferences = memberProxy.References.Where(rf => !UsesQualifiedAccess(rf.Context.Parent));

            //NonQualifiedExternalMemberReferences match containing procedure identifier
            var nonQualRefs = nonQualifiedMemberReferences.Where(rf => identifierMatchesExternal.Any(idm => idm.Prototype == rf.ParentScoping));

            conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, nonQualRefs);

            //NonQualifiedExternalMemberReferences match Field or ModuleConstant declaration identifier
            var qmnsWithMatchingFieldOrModuleConstant = identifierMatchesExternal.Where(idm => IsField(idm) || IsModuleConstant(idm)).Select(idm => idm.QualifiedModuleName);
            var matchingRefs = nonQualifiedMemberReferences.Where(rf => qmnsWithMatchingFieldOrModuleConstant.Contains(rf.QualifiedModuleName));

            conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, matchingRefs);

            //NonQualifiedExternalMemberReferences match other nonQualified member reference(s) within a procedure
            var parentScopeGroupsNonQualifiedMemberReferences = nonQualifiedMemberReferences.GroupBy(rf => rf.ParentScoping);
            foreach (var scopedReferencesGroup in parentScopeGroupsNonQualifiedMemberReferences)
            {
                var matchingNonQualifiedReferences = identifierMatchesAllModules
                                                            .SelectMany(d => d.References)
                                                            .Where(rf => rf.QualifiedModuleName == scopedReferencesGroup.Key.QualifiedModuleName
                                                                                && rf.ParentScoping.Equals(scopedReferencesGroup.Key)
                                                                                && !UsesQualifiedAccess(rf.Context.Parent));

                //A Let\Set\Get identifiers can match a pre-existing property so long as parameters are consistent (parameters are checked elsewhere)
                if (memberProxy.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    var matchingPropertyIdentifiersOfDifferentDeclarationType = matchingNonQualifiedReferences.Where(d => d.Declaration.DeclarationType.HasFlag(DeclarationType.Property) && !d.Declaration.DeclarationType.Equals(memberProxy.DeclarationType));
                    matchingNonQualifiedReferences = matchingNonQualifiedReferences.Except(matchingPropertyIdentifiersOfDifferentDeclarationType);
                }

                conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, matchingNonQualifiedReferences);
            }
            return conflicts.Values.Any();
        }

        private bool LocalDeclarationsHaveSameNameAsParentScope(IConflictDetectionDeclarationProxy proxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesInTargetModule, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            //MS-VBAL 5.4.3.1, 5.4.3.2
            //A local variable/constant cannot have the same name as the containing procedure name.
            var procConflicts = identifierMatchesInTargetModule.Where(idm => (IsLocalVariable(idm)
                                                       || IsLocalConstant(idm))
                                                       && idm.ParentDeclaration.Equals(proxy.Prototype)).ToList();

            conflicts = AddConflicts(conflicts, proxy, procConflicts);
            return conflicts.Values.Any();
        }

        private bool TryFindFunctionNameMatchesParameterConflict(IConflictDetectionDeclarationProxy memberProxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesAllModules, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var parameters = identifierMatchesAllModules
                                    .Where(d => d.DeclarationType.HasFlag(DeclarationType.Parameter)
                                                    && (memberProxy.Prototype?.Equals(d.ParentDeclaration) ?? false));
            conflicts = AddConflicts(conflicts, memberProxy, parameters);
            return conflicts.Values.Any();
        }
    }
}
