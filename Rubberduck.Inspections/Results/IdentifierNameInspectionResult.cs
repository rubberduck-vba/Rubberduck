using System.Globalization;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNameInspectionResult : InspectionResultBase
    {
        public IdentifierNameInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target) {}

        public override string Description
        {
            get { return string.Format(InspectionsUI.IdentifierNameInspectionResultFormat, RubberduckUI.ResourceManager.GetString("DeclarationType_" + Target.DeclarationType, CultureInfo.CurrentUICulture), Target.IdentifierName).Capitalize(); }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
