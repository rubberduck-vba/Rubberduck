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
    /// Finds identifier conflicts for Enumerations and its EnumerationMembers
    /// </summary>
    /// <remarks>
    /// MS-VBAL 5.2.3.4 Enum Declarations
    /// </remarks>
    public class ConflictFinderEnum : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderEnum(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
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

            var destinationModuleDeclarations = GetTargetModuleMembers(enumerationProxy);
            if (!destinationModuleDeclarations.Any())
            {
                return false;
            }

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
            if (!destinationModuleDeclarations.Any())
            {
                return false;
            }

            var identifierMatchingDeclarations
                            = destinationModuleDeclarations.Where(d => d != enumerationMemberProxy.Prototype
                                                                && AreVBAEquivalent(d.IdentifierName, enumerationMemberProxy.IdentifierName))
                                                            .Select(d => sessionData.CreateProxy(d));

            if (ModuleLevelElementChecks(identifierMatchingDeclarations, out var nameConflicts))
            {
                conflicts = AddConflicts(conflicts, enumerationMemberProxy, nameConflicts);
            }
            return conflicts.Values.Any();
        }
    }
}
