using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public interface IRelocateConflictDetection : IConflictDetectionBase
    {
        bool HasConflictInNewLocation(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData);
    }

    public class RelocateConflictDetection : ConflictDetectionBase, IRelocateConflictDetection
    {
        public RelocateConflictDetection(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory)
            : base(declarationFinderProvider, conflictFinderFactory)
        {
        }

        public bool HasConflictInNewLocation(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            var analysisResults = new Dictionary<Declaration, (List<Declaration> Conflicts, string NonConflictName)>
            {
                { proxy.Prototype, (new List<Declaration>(), proxy.IdentifierName) }
            };

            var targetModule = _declarationFinderProvider.DeclarationFinder.MatchName(proxy.TargetModuleName)
                                                                           .OfType<ModuleDeclaration>()
                                                                           .Where(m => m.ProjectId == proxy.ProjectId)
                                                                           .SingleOrDefault();

            var targetModuleComponentType = targetModule?.QualifiedModuleName.ComponentType ?? ComponentType.StandardModule;

            TryFindConflicts(proxy, sessionData, out var conflicts);

            foreach (var conflict in conflicts)
            {
                TryCreateConflictFreeIdentifier(conflict.Key, sessionData, targetModule.IdentifierName, targetModuleComponentType, out var okName, proxy.Accessibility);
                if (!analysisResults.TryGetValue(conflict.Key, out var result))
                {
                    analysisResults.Add(conflict.Key, (conflicts[conflict.Key], okName));
                }
                else
                {
                    analysisResults[conflict.Key] = (conflicts[conflict.Key], okName);
                }
            }

            var enumMemberConflicts = new List<Declaration>();
            var enumMembers = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.EnumerationMember)
                .Where(em => em.ParentDeclaration == proxy.Prototype);

            foreach (var enumMember in enumMembers)
            {
                if (analysisResults.TryGetValue(enumMember, out var results))
                {
                    enumMemberConflicts.AddRange(results.Conflicts);
                }
            }

            var nonConflictRenamePairs = new List<(Declaration, string)>();
            foreach (var resolvedProxy in sessionData.ResolvedProxyDeclarations)
            {
                if (!(enumMembers.Contains(resolvedProxy.Prototype) || proxy.Prototype == resolvedProxy.Prototype))
                {
                    continue;
                }
                if (resolvedProxy.Prototype != null && !AreVBAEquivalent(resolvedProxy.IdentifierName, resolvedProxy.Prototype.IdentifierName))
                {
                    nonConflictRenamePairs.Add((resolvedProxy.Prototype, resolvedProxy.IdentifierName));
                }
            }

            return nonConflictRenamePairs.Any();
        }

        private bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<Declaration, List<Declaration>> conflicts)
        {
            conflicts = new Dictionary<Declaration, List<Declaration>>();

            HasNameConflicts(proxy, sessionData, out var proxyConflicts);
            foreach (var proxyConflict in proxyConflicts)
            {
                conflicts.Add(proxyConflict.Key.Prototype, proxyConflict.Value.Select(prxy => prxy.Prototype).ToList());
            }

            if (proxy.Prototype.DeclarationType.Equals(DeclarationType.Enumeration))
            {
                var enumMemberProxies = CreateEnumerationMemberProxies(proxy, sessionData);
                foreach( var enumMemberProxy in enumMemberProxies)
                {
                    HasNameConflicts(enumMemberProxy, sessionData, out proxyConflicts);
                    foreach (var proxyConflict in proxyConflicts)
                    {
                        conflicts.Add(proxyConflict.Key.Prototype, proxyConflict.Value.Select(prxy => prxy.Prototype).ToList());
                    }
                }
            }
            return conflicts.Values.Any();
        }

        private bool TryCreateConflictFreeIdentifier(Declaration target, IConflictDetectionSessionData sessionData, string destinationName, ComponentType destinationType, out string nonConflictName, Accessibility? accessibility = null)
        {
            var destinationModule = _declarationFinderProvider.DeclarationFinder.MatchName(destinationName)
                                                                               .OfType<ModuleDeclaration>()
                                                                               .SingleOrDefault();

            return TryCreateConflictFreeIdentifier(target, sessionData, destinationModule, out nonConflictName, accessibility);
        }

        private bool TryCreateConflictFreeIdentifier(Declaration target, IConflictDetectionSessionData sessionData, ModuleDeclaration destination, out string nonConflictName, Accessibility? accessibility = null)
        {
            nonConflictName = string.Empty;
            var proxy = CreateProxy(target, sessionData, destination, accessibility);

            if (target.DeclarationType.Equals(DeclarationType.Enumeration))
            {
                var enumMemberProxies = CreateEnumerationMemberProxies(proxy, sessionData);
                foreach (var enumMemberProxy in enumMemberProxies)
                {
                    if (TryResolveToConflictFreeIdentifier(enumMemberProxy, sessionData))
                    {
                        sessionData.RegisterResolvedProxyIdentifier(enumMemberProxy);
                    }
                }
            }

            if (TryResolveToConflictFreeIdentifier(proxy, sessionData))
            {
                sessionData.RegisterResolvedProxyIdentifier(proxy);
                nonConflictName = proxy.IdentifierName;
                return true;
            }
            return false;
        }

        private IConflictDetectionDeclarationProxy CreateProxy(Declaration target, IConflictDetectionSessionData sessionData, ModuleDeclaration destination, Accessibility? accessibility)
        {
            var proxy = sessionData.CreateProxy(target);
            proxy.TargetModule = destination;
            proxy.Accessibility = accessibility ?? target.Accessibility;
            if (target.DeclarationType.Equals(DeclarationType.Enumeration))
            {
                CreateEnumerationMemberProxies(proxy, sessionData);
            }
            return proxy;
        }

        private IEnumerable<IConflictDetectionDeclarationProxy> CreateEnumerationMemberProxies(IConflictDetectionDeclarationProxy enumerationProxy, IConflictDetectionSessionData sessionData)
        {
            var enumMembers = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.EnumerationMember)
                    .Where(en => en.ParentDeclaration == enumerationProxy.Prototype);
            var enumMemberProxies = new List<IConflictDetectionDeclarationProxy>();

            foreach (var enumMember in enumMembers)
            {
                var enumMemberProxy = sessionData.CreateProxy(enumMember);
                enumMemberProxy.TargetModule = enumerationProxy.TargetModule;
                enumMemberProxies.Add(enumMemberProxy);
            }
            return enumMemberProxies;
        }
    }
}
