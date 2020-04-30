using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    class ConflictFinderNonMembers : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderNonMembers(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            if (!IsExistingTargetModule(proxy, out var targetModule))
            {
                return false;
            }

            var allMatches = IdentifierMatches(proxy, sessionData, out var targetModuleMatches);

            //Is Module Variable/Constant
            if (proxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
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
                        .Where(lrf => targetModule.QualifiedModuleName == lrf.QualifiedModuleName);

                    var refConflicts = localRefs.Where(lrf => AreVBAEquivalent(lrf.IdentifierName, proxy.IdentifierName));
                    conflicts = AddReferenceConflicts(conflicts, sessionData, proxy, refConflicts);
                }

                //The nonMember will have Public Accessiblity after the move
                if (proxy.Accessibility != Accessibility.Private && allMatches.Any(id => id.Accessibility != Accessibility.Private && !id.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member)))
                {
                    var publicScopeMatches = allMatches
                                        .Where(match => match.Accessibility != Accessibility.Private
                                                    && match.QualifiedModuleName.HasValue && !AllReferencesUseModuleQualification(match.Prototype, match.QualifiedModuleName.Value));

                    conflicts = AddConflicts(conflicts, proxy, publicScopeMatches);
                }
                return conflicts.Values.Any();
            }

            //Is local Variable/Constant
            var idRefs = proxy.Prototype != null
                                ? _declarationFinderProvider.DeclarationFinder.IdentifierReferences(proxy.Prototype.ParentDeclaration.QualifiedName)
                                : Enumerable.Empty<IdentifierReference>();

            var idRefConflicts = idRefs.Where(rf => AreVBAEquivalent(rf.IdentifierName, proxy.IdentifierName));
            conflicts = AddReferenceConflicts(conflicts, sessionData, proxy, idRefConflicts);

            var localConflicts = targetModuleMatches.Where(idm => idm.ParentDeclaration.Equals(proxy.ParentDeclaration)
                                                    && (IsLocalVariable(idm) || IsLocalConstant(idm)));
            conflicts = AddConflicts(conflicts, proxy, localConflicts);
            return conflicts.Values.Any();
        }

        private bool AllReferencesUseModuleQualification(Declaration declaration, QualifiedModuleName declarationQMN)
        {
            var referencesToModuleQualify = declaration.References.Where(rf => (rf.QualifiedModuleName != declarationQMN));

            return referencesToModuleQualify.All(rf => UsesQualifiedAccess(rf.Context.Parent));
        }
    }
}
