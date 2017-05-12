using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    class DeclarationInspectionResult : InspectionResultBase
    {
        public DeclarationInspectionResult(IInspection inspection, string description, Declaration target, QualifiedContext context = null, Dictionary<string, string> properties = null) :
            base(inspection,
                 description,
                 context == null ? target.QualifiedName.QualifiedModuleName : context.ModuleName,
                 context == null ? target.Context : context.Context,
                 target,
                 target.QualifiedSelection,
                 GetQualifiedMemberName(target),
                 properties)
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
