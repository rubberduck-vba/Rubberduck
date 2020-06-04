using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Fields and Module Constants
    /// </summary>
    class ConflictFinderNonMembers : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderNonMembers(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
            : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var allMatches = IdentifierMatches(proxy, sessionData, out var targetModuleMatches);

            //Is Module Variable/Constant
            if (HasParentWithDeclarationType(proxy, DeclarationType.Module))
            {
                if (ModuleLevelElementChecks(targetModuleMatches, out var moduleLevelConflicts))
                {
                    conflicts = AddConflicts(conflicts, proxy, moduleLevelConflicts);
                }

                var parentScopedGroups = proxy.References.GroupBy(rf => rf.ParentScoping);
                foreach (var scopedReferencesGroup in parentScopedGroups)
                {
                    if (!scopedReferencesGroup.Key.IsMember())
                    {
                        continue;
                    }

                    var localRefs = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(scopedReferencesGroup.Key.QualifiedName)
                        .Where(lrf => lrf.QualifiedModuleName == proxy.TargetModule?.QualifiedModuleName );

                    var refConflicts = localRefs.Where(lrf => AreVBAEquivalent(lrf.IdentifierName, proxy.IdentifierName));
                    conflicts = AddReferenceConflicts(conflicts, sessionData, proxy, refConflicts);
                }

                if (proxy.Accessibility != Accessibility.Private 
                        && allMatches.Any(matchProxy => matchProxy.Accessibility != Accessibility.Private
                        && HasStandardModuleParent(proxy)))
                {

                    var publicScopeMatches = allMatches
                                        .Where(match => match.Accessibility != Accessibility.Private
                                                    && match.Prototype != null 
                                                    && match.QualifiedModuleName.HasValue 
                                                    && !AllReferencesUseModuleQualification(match.Prototype, match.QualifiedModuleName.Value));

                    conflicts = AddConflicts(conflicts, proxy, publicScopeMatches);
                }
                return conflicts.Values.Any();
            }

            //Is local Variable/Constant
            //MS-VBAL 5.4.3.1 and 5.4.3.2
            var idRefs = proxy.Prototype != null
                                ? _declarationFinderProvider.DeclarationFinder.IdentifierReferences(proxy.Prototype.ParentDeclaration.QualifiedName)
                                : Enumerable.Empty<IdentifierReference>();

            var idRefConflicts = idRefs.Where(rf => AreVBAEquivalent(rf.IdentifierName, proxy.IdentifierName));
            conflicts = AddReferenceConflicts(conflicts, sessionData, proxy, idRefConflicts);

            var localConflicts = targetModuleMatches.Where(idm => (idm.ParentDeclaration?.Equals(proxy.ParentDeclaration) ?? false ||  idm.ParentProxy.Equals(proxy.ParentProxy))
                                                    && (IsLocalVariable(idm) || IsLocalConstant(idm) || idm.DeclarationType.Equals(DeclarationType.Parameter)));

            conflicts = AddConflicts(conflicts, proxy, localConflicts);

            var procNameConflicts = targetModuleMatches.Where(idm => (idm.Prototype?.Equals(proxy.ParentDeclaration) ?? false || idm.Equals(proxy.ParentProxy))
                                                    && idm.DeclarationType.HasFlag(DeclarationType.Function));

            conflicts = AddConflicts(conflicts, proxy, procNameConflicts);

            return conflicts.Values.Any();
        }

        private bool AllReferencesUseModuleQualification(Declaration declaration, QualifiedModuleName declarationQMN)
        {
            var referencesToModuleQualify = declaration.References.Where(rf => (rf.QualifiedModuleName != declarationQMN));

            return referencesToModuleQualify.All(rf => UsesQualifiedAccess(rf.Context.Parent));
        }

        private bool HasStandardModuleParent(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.QualifiedModuleName?.ComponentType.Equals(ComponentType.StandardModule)
                ?? proxy.ParentProxy?.ComponentType.Equals(ComponentType.StandardModule)
                ?? false;
        }
        private bool HasMemberParent(IConflictDetectionDeclarationProxy proxy)
        {
            return proxy.ParentProxy?.DeclarationType.HasFlag(DeclarationType.Member)
                ?? proxy.ParentDeclaration?.DeclarationType.HasFlag(DeclarationType.Member)
                ?? false;
        }
    }
}
