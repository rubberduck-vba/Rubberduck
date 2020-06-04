using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Base class for ConflictFinders of entities declared in a Module's Code Section
    /// </summary>
    public abstract class ConflictFinderModuleCodeSection : ConflictFinderBase
    {
        public ConflictFinderModuleCodeSection(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
            : base(declarationFinderProvider, proxyFactory) { }

        public abstract override bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);

        protected bool TryFindMemberConflictChecksCommon(IConflictDetectionDeclarationProxy memberProxy,  IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesInTargetModule, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
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

            if (ReferencesConflictWithProcedureScopeEntities(memberProxy, identifierMatchesInTargetModule, out var referenceConflicts))
            {
                conflicts = AddConflicts(conflicts, memberProxy, referenceConflicts);
            }

            if (memberProxy.DeclarationType.HasFlag(DeclarationType.Function)
                    && TryFindFunctionNameMatchesParameterConflict(memberProxy, identifierMatchesInTargetModule, out var parameterConflicts))
            {
                conflicts = AddConflicts(conflicts, memberProxy, parameterConflicts);
            }
            return conflicts.Values.Any();
        }

        protected bool NonModuleQualifiedReferenceConflicts(IConflictDetectionDeclarationProxy memberProxy, IConflictDetectionSessionData sessionData, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesAllModules, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var identifierMatchesExternal = identifierMatchesAllModules.Where(d => d.QualifiedModuleName != memberProxy.TargetModule?.QualifiedModuleName);

            if (identifierMatchesExternal.Any())
            {
                var conflictReferences = Enumerable.Empty<IdentifierReference>();
                if (TryFindProxyIdentifierConflictsWithOtherNonModuleQualifiedReferences(memberProxy, identifierMatchesExternal, out conflictReferences))
                {
                    conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, conflictReferences);
                }

                var nonQualifiedExternalProxyReferences = NonQualifiedReferences(memberProxy);
                if (nonQualifiedExternalProxyReferences.Any())
                {
                    if (TryFindProxyReferenceConflictsWithOtherModuleDeclarations(memberProxy, identifierMatchesExternal, out conflictReferences))
                    {
                        conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, conflictReferences);
                    }

                    //External NonQualifiedProxyReferences conflicts within Member scope
                    var localEntityIdentifierMatch = identifierMatchesExternal.Where(d => d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member));
                    var localEntityConflictRefs = nonQualifiedExternalProxyReferences.Where(rf => localEntityIdentifierMatch.Select(p => p.ParentDeclaration).Contains(rf.ParentScoping));
                    conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, localEntityConflictRefs);
                }

            }

            return conflicts.Values.Any();
        }

        private bool TryFindProxyIdentifierConflictsWithOtherNonModuleQualifiedReferences(IConflictDetectionDeclarationProxy memberProxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesExternal, out IEnumerable<IdentifierReference> references)
        {
            //Properties have special identifier matching criteria and is handled ConflictFinderProperties
            var matchingExternalEntities = identifierMatchesExternal.Where(idm => !idm.DeclarationType.HasFlag(DeclarationType.Property) && (IsField(idm) || IsModuleConstant(idm) || IsMember(idm)));
            references = matchingExternalEntities.SelectMany(d => d.References)
                                                    .Where(rf => !UsesQualifiedAccess(rf.Context.Parent) 
                                                                && rf.QualifiedModuleName != rf.Declaration.QualifiedModuleName);

            return references.Any();
        }

        private bool TryFindProxyReferenceConflictsWithOtherModuleDeclarations(IConflictDetectionDeclarationProxy memberProxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesExternal, out IEnumerable<IdentifierReference> references)
        {
            var matchingExternalModuleScopeDeclarations = identifierMatchesExternal.Where(d => IsField(d) || IsModuleConstant(d) || IsMember(d));

            var qmnsWithMatchingModuleLevelDeclarations = matchingExternalModuleScopeDeclarations.Select(idm => idm.QualifiedModuleName);
            references = NonQualifiedReferences(memberProxy).Where(rf => qmnsWithMatchingModuleLevelDeclarations.Contains(rf.QualifiedModuleName));

            return references.Any();
        }

        protected bool TryFindDeclarationConflictWithOtherNonModuleQualifiedReferences(IConflictDetectionDeclarationProxy memberProxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatches, out IEnumerable<IdentifierReference> references)
        {
            var identifierMatchesExternal = identifierMatches.Where(d => d.QualifiedModuleName != memberProxy.TargetModule?.QualifiedModuleName);
            var matchingExternalEntities = identifierMatchesExternal.Where(idm => !idm.DeclarationType.HasFlag(DeclarationType.Property) && (IsField(idm) || IsModuleConstant(idm) || IsMember(idm)));
            var matchingReferencesWithinMemberProxyModule = matchingExternalEntities.SelectMany(d => d.References).Where(rf => rf.QualifiedModuleName == memberProxy.QualifiedModuleName);
            references = matchingReferencesWithinMemberProxyModule.Where(rf => !UsesQualifiedAccess(rf.Context.Parent));

            return references.Any();
        }

        private IEnumerable<IdentifierReference> NonQualifiedReferences(IConflictDetectionDeclarationProxy proxy)
        {
            return  proxy.References.Where(rf => !UsesQualifiedAccess(rf.Context.Parent) && rf.QualifiedModuleName != rf.Declaration.QualifiedModuleName);
        }

        private bool LocalDeclarationsHaveSameNameAsParentScope(IConflictDetectionDeclarationProxy proxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesInTargetModule, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            //MS-VBAL 5.4.3.1, 5.4.3.2
            //A local variable/constant cannot have the same name as the containing procedure name.
            var procConflicts = identifierMatchesInTargetModule.Where(idm => (IsLocalVariable(idm)
                                                       || IsLocalConstant(idm)
                                                       || idm.DeclarationType.HasFlag(DeclarationType.Parameter) && proxy.DeclarationType.HasFlag(DeclarationType.Function))
                                                       && idm.ParentDeclaration.Equals(proxy.Prototype)).ToList();

            conflicts = AddConflicts(conflicts, proxy, procConflicts);
            return conflicts.Values.Any();
        }

        private bool ReferencesConflictWithProcedureScopeEntities(IConflictDetectionDeclarationProxy proxy, IEnumerable<IConflictDetectionDeclarationProxy> identifierMatchesInTargetModule, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var localReferenceConflicts = identifierMatchesInTargetModule.Where(idm => (IsLocalVariable(idm)
                                                       || IsLocalConstant(idm)
                                                       || idm.DeclarationType.HasFlag(DeclarationType.Parameter))
                                                            && proxy.References.Any(rf => rf.ParentScoping.Equals(idm.ParentDeclaration)));

            conflicts = AddConflicts(conflicts, proxy, localReferenceConflicts);
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
