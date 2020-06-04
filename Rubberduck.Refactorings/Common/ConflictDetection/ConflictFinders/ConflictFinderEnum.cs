using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Enumerations and EnumerationMembers
    /// </summary>
    /// <remarks>
    /// MS-VBAL 5.2.3.4 Enum Declarations
    /// </remarks>
    public class ConflictFinderEnum : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderEnum(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
            : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            if (proxy.DeclarationType.HasFlag(DeclarationType.Enumeration))
            {
                return TryFindEnumerationNameConflict(proxy, sessionData, out conflicts);
            }
            return TryFindEnumerationMemberNameConflict(proxy, sessionData, out conflicts);
        }

        private bool TryFindEnumerationNameConflict(IConflictDetectionDeclarationProxy enumerationProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            //UserDefinedType and Enumerations have the same conflict rules
            if (UdtAndEnumerationConflicts(enumerationProxy, sessionData, out var udtOrEnumConflicts))
            {
                conflicts = AddConflicts(conflicts, enumerationProxy, udtOrEnumConflicts);
            }

            return conflicts.Values.Any();
        }

        private bool TryFindEnumerationMemberNameConflict(IConflictDetectionDeclarationProxy enumerationMemberProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var destinationModuleDeclarations = GetTargetModuleMembers(enumerationMemberProxy);

            var identifierMatchingDeclarations
                            = destinationModuleDeclarations.Where(d => d != enumerationMemberProxy.Prototype
                                                                && AreVBAEquivalent(d.IdentifierName, enumerationMemberProxy.IdentifierName))
                                                                .Select(d => CreateProxy(sessionData, d));

            var proxyIdentifierMatches = sessionData.RegisteredProxies.Where(rp => AreVBAEquivalent(rp.IdentifierName, enumerationMemberProxy.IdentifierName));

            if (ModuleLevelElementChecks(identifierMatchingDeclarations.Concat(proxyIdentifierMatches), out var nameConflicts))
            {
                var sameEnumMemberNameInOtherEnumDeclarations = nameConflicts.Where(nc => nc.DeclarationType.HasFlag(DeclarationType.EnumerationMember) && nc.ParentDeclaration != enumerationMemberProxy.ParentDeclaration);
                nameConflicts = nameConflicts.Except(sameEnumMemberNameInOtherEnumDeclarations).ToList();
                conflicts = AddConflicts(conflicts, enumerationMemberProxy, nameConflicts);
            }
            return conflicts.Values.Any();
        }

    }
}
