using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionSessionData
    {
        IConflictDetectionDeclarationProxy CreateProxy(Declaration target);
        IConflictDetectionDeclarationProxy CreateProxy(Declaration target, string destinationModuleName, Accessibility? accessibility = null);
        IConflictDetectionDeclarationProxy CreateProxy(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration module, Declaration parentDeclaration, out int retrievalKey);
        void AddProxy(IConflictDetectionDeclarationProxy proxy);
        void RemoveProxy(IConflictDetectionDeclarationProxy proxy);
        IReadOnlyList<IConflictDetectionDeclarationProxy> ResolvedProxyDeclarations { get; }
        void RegisterResolvedProxyIdentifier(IConflictDetectionDeclarationProxy proxy);
        IConflictDetectionDeclarationProxy this[Declaration declaration] { get; }
    }

    public class ConflictDetectionSessionData : IConflictDetectionSessionData
    {
        private Dictionary<int, IConflictDetectionDeclarationProxy> _newDeclarationProxies;
        private Dictionary<Declaration, IConflictDetectionDeclarationProxy> _proxies;
        private List<IConflictDetectionDeclarationProxy> _resolvedProxyDeclarations;

        private readonly IConflictDetectionDeclarationProxyFactory _declarationProxyFactory;

        public ConflictDetectionSessionData(IConflictDetectionDeclarationProxyFactory proxyfactory)
        {
            _resolvedProxyDeclarations = new List<IConflictDetectionDeclarationProxy>();
            _proxies = new Dictionary<Declaration, IConflictDetectionDeclarationProxy>();
            _newDeclarationProxies = new Dictionary<int, IConflictDetectionDeclarationProxy>();
            _declarationProxyFactory = proxyfactory;
        }

        public IReadOnlyList<IConflictDetectionDeclarationProxy> ResolvedProxyDeclarations 
                                                                    => _resolvedProxyDeclarations;

        public void RegisterResolvedProxyIdentifier(IConflictDetectionDeclarationProxy proxy)
        {
            if (!_resolvedProxyDeclarations.Contains(proxy))
            {
                _resolvedProxyDeclarations.Add(proxy);
            }
        }
        public IConflictDetectionDeclarationProxy this[Declaration declaration] => _proxies[declaration];

        public IConflictDetectionDeclarationProxy CreateProxy(Declaration target)
        {
            if (!_proxies.TryGetValue(target, out var proxy))
            {
                proxy = _declarationProxyFactory.Create(target);
                AddProxy(proxy);
            }
            return proxy;
        }

        public IConflictDetectionDeclarationProxy CreateProxy(Declaration target, string destinationModuleName, Accessibility? accessibility = null)
        {
            var proxy = _declarationProxyFactory.Create(target, destinationModuleName);
            proxy.Accessibility = accessibility ?? target.Accessibility;
            if (!_proxies.ContainsKey(proxy.Prototype))
            {
                AddProxy(proxy);
            }
            return proxy;
        }

        public IConflictDetectionDeclarationProxy CreateProxy(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration module, Declaration parentDeclaration, out int retrievalKey)
        {
            var proxy = _declarationProxyFactory.Create(name, declarationType, accessibility, module, parentDeclaration);
            retrievalKey = proxy.GetHashCode();
            if (! _newDeclarationProxies.ContainsKey(retrievalKey))
            {
                AddProxy(proxy);
            }
            return _newDeclarationProxies[retrievalKey];
        }

        public void RemoveProxy(IConflictDetectionDeclarationProxy proxy)
        {
            if (proxy.Prototype != null)
            {
                _proxies.Remove(proxy.Prototype);
            }
            else
            {
                _newDeclarationProxies.Remove(proxy.GetHashCode());
            }
        }

        public void AddProxy(IConflictDetectionDeclarationProxy proxy)
        {
            if (proxy.Prototype != null)
            {
                if (!_proxies.TryGetValue(proxy.Prototype, out _))
                {
                    _proxies.Add(proxy.Prototype, proxy);
                }
                return;
            }

            if (!_newDeclarationProxies.TryGetValue(proxy.GetHashCode(), out _))
            {
                _newDeclarationProxies.Add(proxy.GetHashCode(), proxy);
            }
        }
    }
}
