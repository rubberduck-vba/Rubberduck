using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IRelocateConflictDetection : IConflictDetectionBase
    { }

    public class RelocateConflictDetection : ConflictDetectionBase, IRelocateConflictDetection
    {
        public RelocateConflictDetection(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory)
            : base(declarationFinderProvider, conflictFinderFactory) { }

        public override bool HasConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            if (proxy.DeclarationType.HasFlag(DeclarationType.Enumeration))
            {
                //Relocating an Enumeration relocates its EnumerationMembers - which can also
                //introduce conflicts.  Evaluate this proxy along with its children.
                return !CanResolveEnumerationAndMemberstoConflictFreeIdentifiers(proxy, sessionData);
            }

            return !CanResolveToConflictFreeIdentifier(proxy, sessionData);
        }

        public bool CanResolveEnumerationAndMemberstoConflictFreeIdentifiers(IConflictDetectionDeclarationProxy enumProxy, IConflictDetectionSessionData sessionData)
        {
            var isConflictFree = true;
            var enumMemberProxies = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.EnumerationMember)
                .Where(en => en.ParentDeclaration == enumProxy.Prototype)
                .Select(enm => CreateAndConfigureEnumMember(enm, enumProxy, sessionData));

            var allEnumerationProxies = new List<IConflictDetectionDeclarationProxy>() { enumProxy };
            allEnumerationProxies.AddRange(enumMemberProxies);

            foreach (var proxy in allEnumerationProxies)
            {
                if (!CanResolveToConflictFreeIdentifier(proxy, sessionData))
                {
                    isConflictFree = false;
                }
            }

            return isConflictFree;
        }

        private IConflictDetectionDeclarationProxy CreateAndConfigureEnumMember(Declaration enumMember, IConflictDetectionDeclarationProxy parent, IConflictDetectionSessionData sessionData)
        {
            var proxy = sessionData.CreateProxy(enumMember);
            proxy.TargetModule = parent.TargetModule;
            proxy.IsMutableIdentifier = parent.IsMutableIdentifier;
            return proxy;
        }
    }
}
