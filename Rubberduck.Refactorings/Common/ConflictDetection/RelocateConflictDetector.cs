using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IRelocateConflictDetector
    {
        bool IsConflictingName(Declaration target, ModuleDeclaration destinationModule, out List<(Declaration target, string nonConflictName)> renamePairs, Accessibility? accessibility = null);
    }

    public class RelocateConflictDetector : ConflictDetectorBase, IRelocateConflictDetector
    {
        public RelocateConflictDetector(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory, IDeclarationProxyFactory proxyFactory, IConflictDetectionSessionData session)
            : base(declarationFinderProvider, conflictFinderFactory, proxyFactory, session) { }

        public bool IsConflictingName(Declaration target, ModuleDeclaration destinationModule, out List<(Declaration target, string nonConflictName)> renamePairs, Accessibility? accessibility = null)
        {
            var proxy = CreateProxy(target, destinationModule, accessibility);
            return ProxyRequiresRename(proxy, out renamePairs, accessibility);
        }

        private bool ProxyRequiresRename(IConflictDetectionDeclarationProxy proxy, out List<(Declaration target, string nonConflictName)> renamePairs, Accessibility? accessibility = null)
        {
            renamePairs = new List<(Declaration target, string nonConflictName)>();
            if (proxy.DeclarationType.HasFlag(DeclarationType.Enumeration))
            {
                //Relocating an Enumeration relocates its EnumerationMembers - which can also
                //introduce conflicts.  Evaluate this proxy along with its children.
                ResolveEnumerationAndMemberstoConflictFreeIdentifiers(proxy, out var memberProxies);
                if (!proxy.IdentifierName.Equals(proxy.Prototype.IdentifierName))
                {
                    renamePairs.Add((proxy.Prototype, proxy.IdentifierName));
                }
                foreach (var memberProxy in memberProxies)
                {
                    if (!memberProxy.IdentifierName.Equals(memberProxy.Prototype.IdentifierName))
                    {
                        renamePairs.Add((memberProxy.Prototype, memberProxy.IdentifierName));
                    }
                }
            }
            else
            {
                AssignConflictFreeIdentifier(proxy);
                if (!proxy.IdentifierName.Equals(proxy.Prototype.IdentifierName))
                {
                    renamePairs.Add((proxy.Prototype, proxy.IdentifierName));
                }
            }

            return renamePairs.Any();
        }

        public void ResolveEnumerationAndMemberstoConflictFreeIdentifiers(IConflictDetectionDeclarationProxy enumProxy, out List<IConflictDetectionDeclarationProxy> memberProxies)
        {
            memberProxies = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.EnumerationMember)
                .Where(en => en.ParentDeclaration == enumProxy.Prototype)
                .Select(enm => CreateAndConfigureEnumMember(enm, enumProxy)).ToList();

            var allEnumerationProxies = new List<IConflictDetectionDeclarationProxy>() { enumProxy };
            allEnumerationProxies.AddRange(memberProxies);

            foreach (var proxy in allEnumerationProxies)
            {
                AssignConflictFreeIdentifier(proxy);
            }
        }

        private IConflictDetectionDeclarationProxy CreateAndConfigureEnumMember(Declaration enumMember, IConflictDetectionDeclarationProxy enumeration) 
                    => CreateProxy(enumMember, enumeration.TargetModule);

        private IConflictDetectionDeclarationProxy CreateProxy(Declaration target, ModuleDeclaration module, Accessibility? accessibility = null)
        {
            var proxy = CreateProxy(target);
            proxy.TargetModule = module;
            proxy.Accessibility = accessibility ?? target.Accessibility;
            return proxy;
        }
    }
}
