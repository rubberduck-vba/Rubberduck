using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictFinder
    {
        bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);
    }

    public abstract class ConflictFinderBase : IConflictFinder
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ConflictFinderBase(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public abstract bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);

        protected IEnumerable<IConflictDetectionDeclarationProxy> CreateProxies(IConflictDetectionSessionData sessionData, IEnumerable<Declaration> declarations)
        {
            var proxies = new List<IConflictDetectionDeclarationProxy>();
            foreach (var declaration in declarations)
            {
                var proxy = sessionData.CreateProxy(declaration);
                proxies.Add(proxy);
                if (declaration.DeclarationType.Equals(DeclarationType.EnumerationMember))
                {
                    sessionData[declaration].TargetModule = sessionData[declaration.ParentDeclaration].TargetModule;
                }
            }
            return proxies;
        }

        public bool ModuleLevelElementChecks(IEnumerable<IConflictDetectionDeclarationProxy> matchingDeclarations, out List<IConflictDetectionDeclarationProxy> conflicts)
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
                var idRefProxy = sessionData.CreateProxy(idRef.Declaration);
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
                            .Where(match => match.ProjectId == proxy.ProjectId && match != proxy.Prototype);

            var matchesAllModules = new List<IConflictDetectionDeclarationProxy>();
            foreach (var match in declarationMatches)
            {
                var matchedDeclarationProxy = sessionData.CreateProxy(match);
                matchesAllModules.Add(matchedDeclarationProxy);
            }

            var matchesAllProxyModules = sessionData.ResolvedProxyDeclarations
                            .Where(rp => AreVBAEquivalent(rp.IdentifierName, proxy.IdentifierName)
                                                                && rp.ProjectId == proxy.ProjectId && rp.Prototype != proxy.Prototype);

            matchesAllProxyModules = matchesAllProxyModules.Concat(matchesAllModules);

            targetModuleMatches = matchesAllProxyModules.Where(mod => mod.TargetModuleName == proxy.TargetModuleName)
                            .Where(match => match.ProjectId == proxy.ProjectId && match.Prototype != proxy.Prototype);

            return matchesAllProxyModules;
        }

        protected IEnumerable<IConflictDetectionDeclarationProxy> TargetModuleMatches(IConflictDetectionDeclarationProxy proxy, IEnumerable<IConflictDetectionDeclarationProxy> matchesAllProxyModules)
        {
            return matchesAllProxyModules.Where(mod => mod.TargetModuleName == proxy.TargetModuleName)
                            .Where(match => match.ProjectId == proxy.ProjectId && match.Prototype != proxy.Prototype);
        }

        protected bool IsExistingTargetModule(IConflictDetectionDeclarationProxy proxy, out Declaration targetModule)
        {
            targetModule = null;
            var modules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                .Where(mod => mod.ProjectId == proxy.ProjectId && mod.IdentifierName == proxy.TargetModuleName);
            if (modules.Any())
            {
                targetModule = modules.Single();
                return true;
            }
            return false;
        }

        protected IEnumerable<Declaration> MatchingIdentifierDeclarations(IConflictDetectionDeclarationProxy proxy, params DeclarationType[] declarationTypes)
        {
            var allMatches = new List<Declaration>();
            foreach (var declarationType in declarationTypes)
            {
                var matches = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Where(d => AreVBAEquivalent(d.IdentifierName, proxy.IdentifierName) && d.ProjectId == proxy.ProjectId);

                allMatches.AddRange(matches);
            }
            return allMatches;
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
                && !proxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected bool IsLocalVariable(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Variable)
                && proxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected bool IsModuleConstant(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Constant)
                && !proxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected bool IsLocalConstant(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.DeclarationType.HasFlag(DeclarationType.Constant)
                && proxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
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
