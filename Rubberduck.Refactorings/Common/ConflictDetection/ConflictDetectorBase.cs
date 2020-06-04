using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public class ConflictDetectorBase
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IConflictFinderFactory _conflictFinderFactory;
        private readonly IConflictDetectionSessionData _sessionData;
        protected readonly IDeclarationProxyFactory _proxyFactory;

        public ConflictDetectorBase(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory, IDeclarationProxyFactory proxyFactory, IConflictDetectionSessionData sessionData)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _conflictFinderFactory = conflictFinderFactory;
            _proxyFactory = proxyFactory;
            _sessionData = sessionData;
        }

        protected IConflictDetectionSessionData SessionData => _sessionData;

        public static Func<string, string> ConflictingNameModifier { set; get; } = IncrementIdentifier;

        protected void AssignConflictFreeIdentifier(IConflictDetectionDeclarationProxy proxy, params string[] blackList)
        {
            var isConflictFree = false;
            var iterationMax = 100;
            for (var iteration = 0; iteration < iterationMax && !isConflictFree; iteration++)
            {
                if (blackList.Any(bl => AreVBAEquivalent(bl, proxy.IdentifierName)))
                {
                    proxy.IdentifierName = ConflictingNameModifier(proxy.IdentifierName);
                    continue;
                }
                if (HasNameConflicts(proxy, SessionData, out var conflicts))
                {
                    foreach (var conflict in conflicts)
                    {
                        proxy.IdentifierName = ConflictingNameModifier(proxy.IdentifierName);
                    }
                    continue;
                }
                isConflictFree = true;
            }
        }

        protected bool HasNameConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            if (!IsPotentialProjectNameConflictType(proxy.DeclarationType)
                && !IdentifierIsUsedElsewhereInProject(proxy, sessionData))
            {
                return false;
            }

            var conflictFinder = _conflictFinderFactory.Create(proxy.DeclarationType);

            return conflictFinder.TryFindConflicts(proxy, sessionData, out conflicts);
        }

        public bool IdentifierIsUsedElsewhereInProject(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
            => IdentifierIsUsedElsewhereInProject(proxy.IdentifierName, proxy.ProjectId)
            || sessionData.RegisteredProxies.Any(pxy => AreVBAEquivalent(proxy.IdentifierName, pxy.IdentifierName));

        protected bool IsPotentialProjectNameConflictType(DeclarationType declarationType)
        {
            return declarationType.HasFlag(DeclarationType.Enumeration)
                || declarationType.HasFlag(DeclarationType.UserDefinedType)
                || declarationType.HasFlag(DeclarationType.Project);
        }

        public static string IncrementIdentifier(string identifier)
        {
            var numeric = string.Concat(identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }

        protected IConflictDetectionDeclarationProxy CreateProxy(Declaration target)
        {
            if (!SessionData.TryGetProxyForDeclaration(target, out var proxy))
            {
                proxy = _proxyFactory.CreateProxy(target);
                SessionData.AddProxy(proxy);
            }
            return proxy;
        }

        protected bool AreVBAEquivalent(string idFirst, string idSecond)
            => idFirst.Equals(idSecond, StringComparison.InvariantCultureIgnoreCase);

        private bool IdentifierIsUsedElsewhereInProject(string identifier, string projectID)
            => _declarationFinderProvider.DeclarationFinder.MatchName(identifier)
                            .Except(SessionData.IgnoredDeclarations)
                            .Any(matchedName => matchedName.ProjectId == projectID);
    }
}
