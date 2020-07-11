using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IRenameConflictDetector
    {
        bool IsConflictingName(Declaration target, string newName, out string nonConflictName);
        bool TryFindConflictingDeclarations(Declaration target, string newName, out IEnumerable<Declaration> conflicts);
    }

    public class RenameConflictDetector : ConflictDetectorBase, IRenameConflictDetector
    {
        public RenameConflictDetector(IDeclarationFinderProvider declarationFinderProvider, 
                                        IConflictFinderFactory conflictFinderFactory, 
                                        IDeclarationProxyFactory proxyFactory,
                                        IConflictDetectionSessionData sessionData)
            : base(declarationFinderProvider, conflictFinderFactory, proxyFactory, sessionData)
        {}

        public bool IsConflictingName(Declaration target, string newName, out string nonConflictName)
        {
            return IsConflictingName(CreateProxy(target), newName, out nonConflictName);
        }

        public bool TryFindConflictingDeclarations(Declaration target, string newName, out IEnumerable<Declaration> conflicts)
        {
            conflicts = Enumerable.Empty<Declaration>();
            var proxy = CreateProxy(target);
            if (AreVBAEquivalent(proxy.Prototype.IdentifierName, newName))
            {
                //Detect attempt to change casing
                if (!proxy.Prototype.IdentifierName.Equals(newName, StringComparison.InvariantCulture))
                {
                    conflicts.Concat(new Declaration[] { proxy.Prototype });
                    return true;
                }
                return false;
            }

            proxy.IdentifierName = newName;
            if (HasNameConflicts(proxy, SessionData, out var proxyConflicts))
            {
                conflicts = proxyConflicts.SelectMany(pc => pc.Value).
                                                Where(p => p.Prototype != null)
                                            .Select(cp => cp.Prototype);
                return true;
            }
            return false;
        }

        private bool IsConflictingName(IConflictDetectionDeclarationProxy proxy, string newName, out string nonConflictName)
        {
            if (AreVBAEquivalent(proxy.Prototype.IdentifierName, newName))
            {
                //Detect attempt to change casing
                if (!proxy.Prototype.IdentifierName.Equals(newName, StringComparison.InvariantCulture))
                {
                    nonConflictName = proxy.Prototype.IdentifierName;
                    return true;
                }
                nonConflictName = newName;
                return false;
            }

            proxy.IdentifierName = newName;

            AssignConflictFreeIdentifier(proxy);

            nonConflictName = proxy.IdentifierName;

            return !nonConflictName.Equals(newName);
        }
    }
}
