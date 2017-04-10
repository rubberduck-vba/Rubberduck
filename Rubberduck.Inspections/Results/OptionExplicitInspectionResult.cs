using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class OptionExplicitInspectionResult : InspectionResultBase
    {
        public OptionExplicitInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.OptionExplicitInspectionResultFormat, QualifiedName.ComponentName); }
        }
    }
}
