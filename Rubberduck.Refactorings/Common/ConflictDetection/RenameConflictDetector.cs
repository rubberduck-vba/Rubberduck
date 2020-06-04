using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;

namespace Rubberduck.Refactorings.Common
{
    public interface IRenameConflictDetector
    {
        bool IsConflictingName(Declaration target, string newName, out string nonConflictName);
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
