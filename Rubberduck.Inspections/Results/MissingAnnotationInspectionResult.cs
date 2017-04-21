using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class MissingAnnotationInspectionResult : InspectionResultBase
    {
        public MissingAnnotationInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, QualifiedMemberName? qualifiedName)
            : base(inspection, qualifiedContext.ModuleName, qualifiedName, qualifiedContext.Context) { }

        public override string Description
        {
            get
            {
                var name = QualifiedMemberName.HasValue ? QualifiedMemberName.Value.MemberName : QualifiedName.Name;
                return string.Format("InspectionsUI.MissingAnnotationInspectionResultFormat", name, Context.GetText());
            }
        }
    }
}