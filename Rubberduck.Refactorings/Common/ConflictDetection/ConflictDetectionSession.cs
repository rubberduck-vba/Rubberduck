using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionProxyCreator
    {
        IConflictDetectionDeclarationProxy Create(Declaration target, string newName = null);
        IConflictDetectionModuleDeclarationProxy Create(QualifiedModuleName qmn, string newName = null);
        IConflictDetectionModuleDeclarationProxy CreateNewModule(string projectID, ComponentType componentType, string identifier);
        IConflictDetectionDeclarationProxy CreateNewEntity(IConflictDetectionDeclarationProxy parentProxy, DeclarationType declarationType, string identifier, Accessibility accessibility = Accessibility.Implicit);
    }

    public interface IConflictSession
    {
        /// <summary>
        /// Read-only collection <c>IConflictDetectionDeclarationProxy</c> that have been registered.
        /// </summary>
        /// <remarks>
        /// Registered <c>IConflictDetectionDeclarationProxy</c> participate in subsequent identifier conflict evaluations.
        /// </remarks>
        IReadOnlyCollection<IConflictDetectionDeclarationProxy> RegisteredProxies { get; }

        /// <summary>
        /// Attempts to register an <c>IConflictDetectionDeclarationProxy</c> using its currently assigned <c>Identifier</c>.
        /// </summary>
        /// <param name="nonConflictName">A non-conflicting name is provided if a conflict is found</param>
        /// <param name="forceNoConflictRegistration">If true, the <c>IConflictDetectionDeclarationProxy</c> will be registered with a generated non-conflict name</param>
        /// <returns>Returns false if a conflict is found.  Will always return true if <paramref name="forceNoConflictRegistration"/> is true</returns>
        bool TryRegister(IConflictDetectionDeclarationProxy proxy, out string nonConflictName, bool forceNoConflictRegistration = false);
        /// <summary>
        /// Attempts to register an <c>IConflictDetectionDeclarationProxy</c> in a new Module using its currently assigned <c>Identifier</c>.
        /// </summary>
        /// <param name="nonConflictName">A non-conflicting name is provided if a conflict is found</param>
        /// <param name="forceNoConflictRegistration">If true, the <c>IConflictDetectionDeclarationProxy</c> will be registered with a generated non-conflict name</param>
        /// <returns>Returns false if a conflict is found.  Will always return true if <paramref name="forceNoConflictRegistration"/> is set to true</returns>
        bool TryRegisterRelocation(IConflictDetectionDeclarationProxy proxy, out string nonConflictName, bool forceNoConflictRegistration = false);

        /// <summary>
        /// Removes an <c>IConflictDetectionDeclarationProxy</c> from participating in subsequent identifier conflict evaluations.
        /// </summary>
        void UnRegister(params IConflictDetectionDeclarationProxy[] proxies);

        /// <summary>
        /// Removes an existing <c>Declaration</c> from subsequent conflict evaluations
        /// </summary>
        /// <remarks>
        /// Registration of an <c>IConflictDetectionDeclarationProxy</c> automatically ignores the proxy's
        /// underlying <c>Declaration</c> 
        /// </remarks>
        void IgnoreConflictDetection(Declaration declaration);
        /// <summary>
        /// Includes an existing <c>Declaration</c> in subsequent conflict evaluations
        /// </summary>
        void RestoreConflictDetection(Declaration declaration);

        /// <summary>
        /// Read-only collection of <c>Declaration</c> that are not evaluated in conflict analysis
        /// </summary>
        IReadOnlyCollection<Declaration> IgnoredDeclarations { get; }

        IRenameConflictDetector RenameConflictDetector { get; }
        IRelocateConflictDetector RelocateConflictDetector { get; }
        INewModuleConflictDetector NewModuleConflictDetector { get; }
        INewEntityConflictDetector NewEntityConflictDetector { get; }

        /// <summary>
        /// ProxyCreator supports creation of new <c>ConflictDetectionDeclarationProxy</c> objects for conflict evaluations 
        /// </summary>
        IConflictDetectionProxyCreator ProxyCreator { get; }

        /// <summary>
        /// Returns a set of (<c>Declaration</c>, <c>string</c>) pairs that have been
        /// registered as a non-conflicting rename of an existing <c>Declaration</c>
        /// </summary>
        /// <remarks>
        /// It is possible for a single<c>ConflictDetectionDeclarationProxy</c> to result in more 
        /// than one rename pair. (e.g. Moving an enumeration may result in multiple enumeration Member conflicts).
        /// </remarks>
        IReadOnlyCollection<(Declaration Target, string NewName)> RenamePairs { get; }

        /// <summary>
        /// Delegate function used to modify a conflicting Identifier
        /// </summary>
        Func<string, string> ConflictNameModifier { set; get; }
    }

    public class ConflictDetectionSession : IConflictSession, IConflictDetectionProxyCreator
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IConflictDetectorFactory _detectorFactory;
        private readonly IDeclarationProxyFactory _proxyFactory;

        private readonly IConflictDetectionSessionData _sessionData;

        public ConflictDetectionSession(IDeclarationFinderProvider declarationFinderProvider,
                                        IDeclarationProxyFactory proxyFactory,
                                        IConflictDetectorFactory detectorFactory,
                                        IConflictDetectionSessionData sessionData)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _proxyFactory = proxyFactory;
            _detectorFactory = detectorFactory;

            _sessionData = sessionData;

            RenameConflictDetector = _detectorFactory.CreateRenameConflictDetector(sessionData);
            RelocateConflictDetector = _detectorFactory.CreateRelocateConflictDetector(sessionData);
            NewModuleConflictDetector = _detectorFactory.CreateNewModuleConflictDetector(sessionData);
            NewEntityConflictDetector = _detectorFactory.CreateNewEntityConflictDetector(sessionData);
        }

        public IRenameConflictDetector RenameConflictDetector { private set; get; }
        public IRelocateConflictDetector RelocateConflictDetector { private set; get; }
        public INewModuleConflictDetector NewModuleConflictDetector { private set; get; }
        public INewEntityConflictDetector NewEntityConflictDetector { private set; get; }

        public IConflictDetectionModuleDeclarationProxy Create(QualifiedModuleName qmn, string newName = null)
        {
            var module = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(qmn);
            return new ModuleConflictDetectionDeclarationProxy(module as ModuleDeclaration);
        }

        public IConflictDetectionDeclarationProxy Create(Declaration target, string newName = null)
        {
            if (target is ModuleDeclaration)
            {
                return Create(target.QualifiedModuleName, newName);
            }

            var proxy = _proxyFactory.CreateProxy(target);
            if (!string.IsNullOrEmpty(newName))
            {
                proxy.IdentifierName = newName;
            }
            _sessionData.AddProxy(proxy);
            return proxy;
        }

        public IConflictDetectionProxyCreator ProxyCreator => this as IConflictDetectionProxyCreator;

        public IConflictDetectionModuleDeclarationProxy CreateNewModule(string projectID, ComponentType componentType, string identifier)
        {
            var proxy = _proxyFactory.CreateProxyNewModule(projectID, componentType, identifier);
            _sessionData.AddProxy(proxy);
            return proxy;
        }

        public IConflictDetectionDeclarationProxy CreateNewEntity(IConflictDetectionDeclarationProxy parentProxy, DeclarationType declarationType, string identifier,Accessibility accessibility = Accessibility.Implicit)
        {
            var proxy = _proxyFactory.CreateNewEntityProxy(identifier, declarationType, accessibility, parentProxy);
            _sessionData.AddProxy(proxy);
            return proxy;
        }

        public IReadOnlyCollection<IConflictDetectionDeclarationProxy> NewDeclarationProxies
            => _sessionData.RegisteredProxies.Where(rp => rp.Prototype == null).ToList();

        public IReadOnlyCollection<IConflictDetectionDeclarationProxy> RegisteredProxies
            => _sessionData.RegisteredProxies.ToList();

        public bool TryRegister(IConflictDetectionDeclarationProxy proxy, out string nonConflictName, bool forceNoConflictRegistration = false)
        {
            nonConflictName = proxy.IdentifierName;
            var parentProxy = proxy.ParentDeclaration is null
                        ? proxy.ParentProxy
                        : Create(proxy.ParentDeclaration);

            if (parentProxy is IConflictDetectionModuleDeclarationProxy moduleProxy)
            {
                var newEntityConflictDetector = NewEntityConflictDetector;

                if (newEntityConflictDetector.HasConflictingName(proxy, out nonConflictName)
                    && !forceNoConflictRegistration)
                {
                    return false;
                }
            }
            else
            {
                var newEntityConflictDetector = NewEntityConflictDetector;
                if (newEntityConflictDetector.HasConflictingName(proxy, out nonConflictName)
                    && !forceNoConflictRegistration)
                {
                    return false;
                }
            }
            proxy.IdentifierName = nonConflictName;
            Register(proxy);
            return true;
        }

        public bool TryRegisterRelocation(IConflictDetectionDeclarationProxy proxy, out string nonConflictName, bool forceNoConflictRegistration = false)
        {
            nonConflictName = proxy.IdentifierName;
            var parentProxy = proxy.ParentDeclaration is null
                        ? proxy.ParentProxy
                        : Create(proxy.ParentDeclaration);

            RelocateConflictDetector.IsConflictingName(proxy.Prototype, proxy.TargetModule, out var renamePairs, proxy.Accessibility);
            if (!forceNoConflictRegistration)
            {
                return false;
            }

            foreach (var renamePair in renamePairs)
            {
                _sessionData.TryGetProxyForDeclaration(renamePair.target, out var declarationProxy);
                Register(declarationProxy);
            }
            return true;
        }

        private void Register(IConflictDetectionDeclarationProxy proxy)
        {
            _sessionData.RegisterProxy(proxy);
            if (proxy.Prototype != null)
            {
                _sessionData.IgnoreConflictDetection(proxy.Prototype);
            }
        }

        public void UnRegister(params IConflictDetectionDeclarationProxy[] proxies)
        {
            foreach (var proxy in proxies)
            {
                _sessionData.UnRegisterProxy(proxy);
                if (proxy.Prototype != null)
                {
                    _sessionData.RestoreConflictDetection(proxy.Prototype);
                }
            }
        }

        public IReadOnlyCollection<Declaration> IgnoredDeclarations 
            => _sessionData.IgnoredDeclarations;

        public void IgnoreConflictDetection(Declaration declaration) 
            => _sessionData.IgnoreConflictDetection(declaration);

        public void RestoreConflictDetection(Declaration declaration) 
            => _sessionData.RestoreConflictDetection(declaration);

        public IReadOnlyCollection<(Declaration Target, string NewName)> RenamePairs
            => _sessionData.RenamePairs;

        public Func<string, string> ConflictNameModifier
        {
            set => ConflictDetectorBase.ConflictingNameModifier = value;
            get => ConflictDetectorBase.ConflictingNameModifier;
        }
    }
}
