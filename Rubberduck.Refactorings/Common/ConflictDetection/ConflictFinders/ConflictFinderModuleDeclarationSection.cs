using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Supporting base class for a Module's Declaration Section
    /// <seealso cref="ConflictFinderNonMembers"/>
    /// <seealso cref="ConflictFinderEnum"/>
    /// <seealso cref="ConflictFinderUDT"/>
    /// </summary>
    public abstract class ConflictFinderModuleDeclarationSection : ConflictFinderBase
    {
        public ConflictFinderModuleDeclarationSection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public abstract override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts);

        protected bool UdtAndEnumerationConflicts(IConflictDetectionDeclarationProxy proxyEntity, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var destinationModuleDeclarations = GetTargetModuleMembers(proxyEntity);

            var udtIdentifierConflictTypes = new List<DeclarationType>()
            {
                DeclarationType.UserDefinedType,
                DeclarationType.Enumeration,
            };

            foreach (var potentialConflict in destinationModuleDeclarations.Where(pc => AreVBAEquivalent(pc.IdentifierName, proxyEntity.IdentifierName)))
            {
                if (udtIdentifierConflictTypes.Any(ect => potentialConflict.DeclarationType.HasFlag(ect)))
                {
                    conflicts = AddConflicts(conflicts, proxyEntity, sessionData.CreateProxy(potentialConflict));
                }
            }

            if (proxyEntity.Accessibility != Accessibility.Private)
            {
                var conflictingModuleIdentifiers = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                    .Where(m => m.ProjectId == proxyEntity.ProjectId
                                                    && AreVBAEquivalent(m.IdentifierName, proxyEntity.IdentifierName));

                conflicts = AddConflicts(conflicts, proxyEntity, CreateProxies(sessionData, conflictingModuleIdentifiers));

                var conflictingProjectIdentifiers = _declarationFinderProvider.DeclarationFinder.Projects
                    .Where(p => AreVBAEquivalent(p.IdentifierName, proxyEntity.IdentifierName));

                conflicts = AddConflicts(conflicts, proxyEntity, CreateProxies(sessionData, conflictingProjectIdentifiers));

                var conflictingUDTsInProject = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.UserDefinedType)
                        .Where(udtCandidate => udtCandidate.ProjectId == proxyEntity.ProjectId
                                                    && udtCandidate != proxyEntity.Prototype
                                                    && !udtCandidate.HasPrivateAccessibility()
                                                    && AreVBAEquivalent(udtCandidate.IdentifierName, proxyEntity.IdentifierName));

                conflicts = AddConflicts(conflicts, proxyEntity, CreateProxies(sessionData, conflictingUDTsInProject));

                var conflictingEnumsInProject = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Enumeration)
                        .Where(enumCandidate => enumCandidate.ProjectId == proxyEntity.ProjectId
                                                    && enumCandidate != proxyEntity.Prototype
                                                    && !enumCandidate.HasPrivateAccessibility()
                                                    && AreVBAEquivalent(enumCandidate.IdentifierName, proxyEntity.IdentifierName));

                conflicts = AddConflicts(conflicts, proxyEntity, CreateProxies(sessionData, conflictingEnumsInProject));
            }
            return conflicts.Values.Any();
        }

        protected IEnumerable<Declaration> GetTargetModuleMembers(IConflictDetectionDeclarationProxy proxy)
        {
            var modules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                .Where(mod => mod.ProjectId == proxy.ProjectId && mod.IdentifierName == proxy.TargetModuleName);
            if (modules.Any())
            {
                return _declarationFinderProvider.DeclarationFinder.Members(modules.Single().QualifiedModuleName);
            }
            return Enumerable.Empty<Declaration>();
        }
    }
}
