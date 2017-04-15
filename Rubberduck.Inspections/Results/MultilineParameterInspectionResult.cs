using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class MultilineParameterInspectionResult : InspectionResultBase
    {
        public MultilineParameterInspectionResult(IInspection inspection, QualifiedContext qualifiedContext, QualifiedMemberName? qualifiedName)
            : base(inspection, qualifiedContext.ModuleName, qualifiedName, qualifiedContext.Context) {}

        public override string Description
        {
            get
            {
                return string.Format(
                    Context.GetSelection().LineCount > 3
                        ? RubberduckUI.EasterEgg_Continuator
                        : InspectionsUI.MultilineParameterInspectionResultFormat, Target.IdentifierName).Capitalize();
            }
        }
    }
}
