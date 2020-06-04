using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionSessionData
    {
        bool TryGetProxyForDeclaration(Declaration declaration, out IConflictDetectionDeclarationProxy proxy);
        void AddProxy(IConflictDetectionDeclarationProxy proxy);
        void RegisterProxy(IConflictDetectionDeclarationProxy proxy);
        void UnRegisterProxy(IConflictDetectionDeclarationProxy proxy);
        IReadOnlyList<IConflictDetectionDeclarationProxy> RegisteredProxies { get; }
        IReadOnlyList<(Declaration Target, string NewName)> RenamePairs { get; }
        void IgnoreConflictDetection(Declaration declaration);
        void RestoreConflictDetection(Declaration declaration);
        IReadOnlyCollection<Declaration> IgnoredDeclarations { get; }
    }

    public class ConflictDetectionSessionData : IConflictDetectionSessionData
    {
        private HashSet<IConflictDetectionDeclarationProxy> _newDeclarationProxies;
        private Dictionary<Declaration, IConflictDetectionDeclarationProxy> _prototypedProxies;
        private HashSet<IConflictDetectionDeclarationProxy> _registeredProxies;
        private Dictionary<Declaration, string> _registeredRenamePairs;
        private HashSet<Declaration> _declarationsToIgnore;

        public ConflictDetectionSessionData()
        {
            _registeredRenamePairs = new Dictionary<Declaration, string>();
            _registeredProxies = new HashSet<IConflictDetectionDeclarationProxy>();
            _prototypedProxies = new Dictionary<Declaration, IConflictDetectionDeclarationProxy>();
            _newDeclarationProxies = new HashSet<IConflictDetectionDeclarationProxy>();
            _declarationsToIgnore = new HashSet<Declaration>();
        }

        public IReadOnlyList<IConflictDetectionDeclarationProxy> RegisteredProxies
                                => _registeredProxies.ToList();

        public IReadOnlyList<(Declaration Target, string NewName)> RenamePairs
                                => _registeredProxies.Where(pxy => pxy.Prototype != null 
                                                                && !pxy.IdentifierName.Equals(pxy.Prototype.IdentifierName, StringComparison.InvariantCultureIgnoreCase))
                                                        .Select(pxy => (pxy.Prototype, pxy.IdentifierName)).ToList();

        public bool TryGetProxyForDeclaration(Declaration declaration, out IConflictDetectionDeclarationProxy proxy)
                                => _prototypedProxies.TryGetValue(declaration, out proxy);

        public void AddProxy(IConflictDetectionDeclarationProxy proxy)
        {
            if (proxy.Prototype is null)
            {
                if (!_newDeclarationProxies.Contains(proxy))
                {
                    _newDeclarationProxies.Add(proxy);
                }
            }
            else if (!_prototypedProxies.ContainsKey(proxy.Prototype))
            {
                _prototypedProxies.Add(proxy.Prototype, proxy);
            }
        }

        public void RegisterProxy(IConflictDetectionDeclarationProxy proxy)
        {
            AddProxy(proxy);
            _registeredProxies.Add(proxy);
        }

        public void UnRegisterProxy(IConflictDetectionDeclarationProxy proxy) 
            => _registeredProxies.Remove(proxy);

        public IReadOnlyCollection<Declaration> IgnoredDeclarations => _declarationsToIgnore;

        public void IgnoreConflictDetection(Declaration declaration) 
            => _declarationsToIgnore.Add(declaration);

        public void RestoreConflictDetection(Declaration declaration) 
            => _declarationsToIgnore.Remove(declaration);
    }
}
