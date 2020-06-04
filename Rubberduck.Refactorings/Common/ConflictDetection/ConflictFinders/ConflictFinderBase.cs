using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictFinder
    {
        bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);
    }

    public abstract class ConflictFinderBase : IConflictFinder
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected readonly IDeclarationProxyFactory _proxyFactory;

        public ConflictFinderBase(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _proxyFactory = proxyFactory;
        }

        public abstract bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);

        protected IEnumerable<IConflictDetectionDeclarationProxy> CreateProxies(IConflictDetectionSessionData sessionData, IEnumerable<Declaration> declarations)
        {
            var proxies = new List<IConflictDetectionDeclarationProxy>();
            foreach (var declaration in declarations)
            {
                var proxy = CreateProxy(sessionData, declaration);

                proxies.Add(proxy);
                if (declaration.DeclarationType.Equals(DeclarationType.EnumerationMember))
                {
                    if (sessionData.TryGetProxyForDeclaration(declaration.ParentDeclaration, out var enumeration))
                    {
                        proxy.TargetModule = enumeration.TargetModule;
                    }
                }
            }
            return proxies;
        }

        protected IConflictDetectionDeclarationProxy CreateProxy(IConflictDetectionSessionData sessionData, Declaration declaration)
        {
            if (!sessionData.TryGetProxyForDeclaration(declaration, out var proxy))
            {
                proxy = _proxyFactory.CreateProxy(declaration);
                sessionData.AddProxy(proxy);
            }
            return proxy;
        }

        protected bool ModuleLevelElementChecks(IEnumerable<IConflictDetectionDeclarationProxy> matchingDeclarations, out List<IConflictDetectionDeclarationProxy> conflicts)
        {
            conflicts = new List<IConflictDetectionDeclarationProxy>();

            foreach (var identifierMatch in matchingDeclarations)
            {
                if (IsMember(identifierMatch)
                    || IsField(identifierMatch)
                    || IsModuleConstant(identifierMatch)
                    || IsEnumMember(identifierMatch))
                {
                    conflicts.Add(identifierMatch);
                }
            }
            return conflicts.Any();
        }

        protected Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> AddReferenceConflicts(Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts, IConflictDetectionSessionData sessionData, IConflictDetectionDeclarationProxy target, IEnumerable<IdentifierReference> conflictReferencesFound)
        {
            if (!conflictReferencesFound.Any())
            {
                return conflicts;
            }

            var refConflicts = new List<IConflictDetectionDeclarationProxy>();
            foreach (var idRef in conflictReferencesFound)
            {
                var idRefProxy = CreateProxy(sessionData, idRef.Declaration);
                refConflicts.Add(idRefProxy);
            }

            if (!conflicts.TryGetValue(target, out var existingConflicts))
            {
                conflicts.Add(target, new List<IConflictDetectionDeclarationProxy>());
            }

            var found = conflicts[target].Concat(refConflicts);

            conflicts[target] = found.Distinct().ToList();
            return conflicts;
        }

        protected Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> AddConflicts(Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts, IConflictDetectionDeclarationProxy target, params IConflictDetectionDeclarationProxy[] conflictsFound)
        {
            return AddConflicts(conflicts, target, conflictsFound.ToList());
        }

        protected Dictionary<T, List<K>> AddConflicts<T, K>(Dictionary<T, List<K>> conflicts, T target, IEnumerable<K> conflictsFound)
        {
            if (!conflictsFound.Any())
            {
                return conflicts;
            }

            if (!conflicts.TryGetValue(target, out var existingConflicts))
            {
                conflicts.Add(target, new List<K>());
            }

            var found = conflicts[target].Concat(conflictsFound);

            conflicts[target] = found.Distinct().ToList();
            return conflicts;
        }

        protected Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> AddConflicts(Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts, IConflictDetectionDeclarationProxy target, IEnumerable<IConflictDetectionDeclarationProxy> conflictsFound)
        {
            if (!conflictsFound.Any())
            {
                return conflicts;
            }

            if (!conflicts.TryGetValue(target, out _))
            {
                conflicts.Add(target, new List<IConflictDetectionDeclarationProxy>());
            }

            conflicts[target] = conflicts[target].Concat(conflictsFound).ToList();
            return conflicts;
        }

        protected Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> AddConflicts(Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts, IConflictDetectionDeclarationProxy target, Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflictsFound)
        {
            if (!conflictsFound.Any())
            {
                return conflicts;
            }

            foreach (var kvPair in conflictsFound)
            {
                conflicts = AddConflicts(conflicts, kvPair.Key, conflictsFound[kvPair.Key]);
            }
            return conflicts;
        }

        protected IEnumerable<IConflictDetectionDeclarationProxy> IdentifierMatches(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out IEnumerable<IConflictDetectionDeclarationProxy> targetModuleMatches)
        {
            var declarationMatches = _declarationFinderProvider.DeclarationFinder.MatchName(proxy.IdentifierName)
                            .Except(sessionData.IgnoredDeclarations)
                            .Where(match => match.ProjectId == proxy.ProjectId
                                        && match != proxy.Prototype);

            var matchesAllModules = new List<IConflictDetectionDeclarationProxy>();
            foreach (var match in declarationMatches)
            {
                if (sessionData.TryGetProxyForDeclaration(match, out var matchProxy) && AreVBAEquivalent(matchProxy.IdentifierName, match.IdentifierName))
                {
                    matchesAllModules.Add(matchProxy);
                    continue;
                }

                if (matchProxy is null)
                {
                    matchesAllModules.Add(CreateProxy(sessionData, match));
                }
            }

            var matchesAllProxyModules = sessionData.RegisteredProxies
                            .Where(rp => AreVBAEquivalent(rp.IdentifierName, proxy.IdentifierName)
                                                                && rp.ProjectId == proxy.ProjectId 
                                                                && (rp.Prototype != proxy.Prototype
                                                                    || rp.Prototype == null && proxy.Prototype == null));

            matchesAllProxyModules = matchesAllProxyModules.Concat(matchesAllModules).Where(m => m.ProxyID != proxy.ProxyID);

            targetModuleMatches = matchesAllProxyModules.Where(mod => mod.TargetModuleName == proxy.TargetModuleName)
                            .Where(match => match.ProjectId == proxy.ProjectId
                                                    && (match.Prototype != proxy.Prototype
                                                            || match.Prototype == null && proxy.Prototype == null));

            return matchesAllProxyModules;
        }

        protected IEnumerable<Declaration> ModuleOrProjectIdentifierConflicts(IConflictDetectionDeclarationProxy proxy)
        {
            Debug.Assert(proxy.DeclarationType.HasFlag(DeclarationType.Project) || proxy.DeclarationType.HasFlag(DeclarationType.Module));

            Predicate<Declaration> projectProjectID = (d) => {return true; };
            Predicate<Declaration> moduleProjectID = (d) => { return d.ProjectId == proxy.ProjectId; };

            //MS-VBAL 5.2.3.3 and 5.2.3.4 - Project name cannot match any Public UDT or Enum identifiers 
            var projectIDPredicate = proxy.DeclarationType.HasFlag(DeclarationType.Project)
                                            ? projectProjectID
                                            : moduleProjectID;

            var matchingModules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                .Where(mod => AreVBAEquivalent(mod.IdentifierName, proxy.IdentifierName) && mod.ProjectId == proxy.ProjectId);

            var matchingUDTs = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.UserDefinedType)
                .Where(udt => AreVBAEquivalent(udt.IdentifierName, proxy.IdentifierName) && udt.Accessibility != Accessibility.Private && projectIDPredicate(udt));


            var matchingEnums = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Enumeration)
                .Where(en => AreVBAEquivalent(en.IdentifierName, proxy.IdentifierName) && en.Accessibility != Accessibility.Private && projectIDPredicate(en));

            return matchingModules.Concat(matchingUDTs).Concat(matchingEnums);
        }

        protected IEnumerable<IConflictDetectionDeclarationProxy> ModuleOrProjectProxyConflicts(IConflictDetectionDeclarationProxy proxy, IEnumerable<IConflictDetectionDeclarationProxy> registeredProxies)
        {
            Debug.Assert(proxy.DeclarationType.HasFlag(DeclarationType.Project) || proxy.DeclarationType.HasFlag(DeclarationType.Module));

            if (!registeredProxies.Any())
            {
                return Enumerable.Empty<IConflictDetectionDeclarationProxy>();
            }

            Predicate<IConflictDetectionDeclarationProxy> projectProjectID = (d) => { return true; };
            Predicate<IConflictDetectionDeclarationProxy> moduleProjectID = (d) => { return d.ProjectId == proxy.ProjectId; };

            //MS-VBAL 5.2.3.3 and 5.2.3.4 - Project name cannot match any Public UDT or Enum identifiers 
            var projectIDPredicate = proxy.DeclarationType.HasFlag(DeclarationType.Project)
                                            ? projectProjectID
                                            : moduleProjectID;

            var matchingEnumProxies = registeredProxies.Where(rp => rp.DeclarationType.HasFlag(DeclarationType.Enumeration)
                  && AreVBAEquivalent(rp.IdentifierName, proxy.IdentifierName) && rp.Accessibility != Accessibility.Private && projectIDPredicate(rp));

            var matchingUdtProxies = registeredProxies.Where(rp => rp.DeclarationType.HasFlag(DeclarationType.UserDefinedType)
                 && AreVBAEquivalent(rp.IdentifierName, proxy.IdentifierName) && rp.Accessibility != Accessibility.Private && projectIDPredicate(rp));

            var matchingModuleProxies = registeredProxies.Where(rp => rp.DeclarationType.HasFlag(DeclarationType.Module)
                                                    && rp.ProjectId.Equals(proxy.ProjectId)
                                                    && AreVBAEquivalent(rp.IdentifierName, proxy.IdentifierName));

            return matchingModuleProxies.Concat(matchingUdtProxies).Concat(matchingEnumProxies).ToList();
        }

        protected bool UsesQualifiedAccess(RuleContext ruleContext)
        {
            return (ruleContext is VBAParser.WithMemberAccessExprContext)
                || (ruleContext is VBAParser.MemberAccessExprContext);
        }

        protected bool IsMember(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected bool IsField(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Variable)
                && !HasParentWithDeclarationType(proxy, DeclarationType.Member);
        }

        protected bool IsLocalVariable(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Variable)
                && HasParentWithDeclarationType(proxy, DeclarationType.Member);
        }

        protected bool IsModuleConstant(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Constant)
                && !HasParentWithDeclarationType(proxy, DeclarationType.Member);
        }

        protected bool IsLocalConstant(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Constant)
                && HasParentWithDeclarationType(proxy, DeclarationType.Member);
        }

        protected bool HasParentWithDeclarationType(IConflictDetectionDeclarationProxy proxy, params  DeclarationType[] declarationTypes)
        {
            foreach (var declarationType in declarationTypes)
            {
                if( proxy.ParentDeclaration?.DeclarationType.HasFlag(declarationType)
                    ?? proxy.ParentProxy.DeclarationType.HasFlag(declarationType))
                {
                    return true;
                }
            }
            return false;
        }

        private bool IsEnumMember(Declaration declaration)
        {
            return declaration.DeclarationType.Equals(DeclarationType.EnumerationMember);
        }

        private bool IsEnumMember(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.Equals(DeclarationType.EnumerationMember);
        }

        protected bool AreVBAEquivalent(string idFirst, string idSecond)
            => idFirst.Equals(idSecond, StringComparison.InvariantCultureIgnoreCase);
    }
}
