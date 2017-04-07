using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public sealed class ObjectVariableNotSetInspectionResult : InspectionResultBase
    {
        public ObjectVariableNotSetInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context, reference.Declaration)
        {
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ObjectVariableNotSetInspectionResultFormat, Target.IdentifierName).Capitalize(); }
        }
    }
}