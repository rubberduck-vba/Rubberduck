using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    class DeclarationInspectionResult : InspectionResultBase
    {
        public DeclarationInspectionResult(IInspection inspection, string description, Declaration target) :
            base(inspection,
                 description,
                 target.QualifiedName.QualifiedModuleName,
                 target.Context,
                 target,
                 target.QualifiedSelection,
                 GetQualifiedMemberName(target))
        {
        }
        
        private static QualifiedMemberName? GetQualifiedMemberName(Declaration target)
        {
            if (string.IsNullOrEmpty(target?.QualifiedName.QualifiedModuleName.ComponentName))
            {
                return null;
            }

            if (target.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return target.QualifiedName;
            }

            return GetQualifiedMemberName(target.ParentDeclaration);
        }
    }
}
