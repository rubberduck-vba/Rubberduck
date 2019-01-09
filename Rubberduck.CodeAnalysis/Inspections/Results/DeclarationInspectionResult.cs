using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class DeclarationInspectionResult : InspectionResultBase
    {
        public DeclarationInspectionResult(IInspection inspection, string description, Declaration target, QualifiedContext context = null, dynamic properties = null) :
            base(inspection,
                 description,
                 context == null ? target.QualifiedName.QualifiedModuleName : context.ModuleName,
                 context == null ? target.Context : context.Context,
                 target,
                 target.QualifiedSelection,
                 GetQualifiedMemberName(target),
                 (object)properties)
        {
        }
        
        private static QualifiedMemberName? GetQualifiedMemberName(Declaration target)
        {
            if (string.IsNullOrEmpty(target?.QualifiedName.QualifiedModuleName.ComponentName))
            {
                return null;
            }

            return target.DeclarationType.HasFlag(DeclarationType.Member)
                ? target.QualifiedName
                : GetQualifiedMemberName(target.ParentDeclaration);
        }

        public override bool ChangesInvalidateResult(ICollection<QualifiedModuleName> modifiedModules)
        {
            return modifiedModules.Contains(Target.QualifiedModuleName)
                   || base.ChangesInvalidateResult(modifiedModules);
        }
    }
}
