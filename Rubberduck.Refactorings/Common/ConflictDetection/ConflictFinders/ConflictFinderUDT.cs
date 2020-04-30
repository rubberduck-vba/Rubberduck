using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public class ConflictFinderUDT : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderUDT(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            if (proxy.DeclarationType.HasFlag(DeclarationType.UserDefinedType))
            {
                return TryFindUDTNameConflict(proxy, sessionData, out conflicts);
            }
            return TryFindUDTMemberNameConflict(proxy, sessionData, out conflicts);
        }


        //MS-VBAL 5.2.3.3 UserDefinedType Declarations
        private bool TryFindUDTNameConflict(IConflictDetectionDeclarationProxy udtProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            //UserDefinedType and Enumerations have the same conflict rules
            if (UdtAndEnumerationConflicts(udtProxy, sessionData, out var udtOrEnumConflicts))
            {
                conflicts = AddConflicts(conflicts, udtProxy, udtOrEnumConflicts);
            }

            return conflicts.Values.Any();
        }

        private bool TryFindUDTMemberNameConflict(IConflictDetectionDeclarationProxy udtMemberProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var relatedMembers = _declarationFinderProvider.DeclarationFinder.Members(udtMemberProxy.TargetModule.QualifiedModuleName)
                                                .Where(d => d.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember) && d.ParentDeclaration == udtMemberProxy.Prototype.ParentDeclaration);
            var memberConflicts = relatedMembers.Where(rm => AreVBAEquivalent(rm.IdentifierName, udtMemberProxy.IdentifierName))
                                                    .Select(d => sessionData.CreateProxy(d));

            conflicts = AddConflicts(conflicts, udtMemberProxy, memberConflicts);
            return conflicts.Values.Any();
        }
    }
}
