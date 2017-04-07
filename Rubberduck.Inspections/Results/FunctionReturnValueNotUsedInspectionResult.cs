using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class FunctionReturnValueNotUsedInspectionResult : InspectionResultBase
    {
        public FunctionReturnValueNotUsedInspectionResult(IInspection inspection, ParserRuleContext context, QualifiedMemberName qualifiedName, Declaration target)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target) {}

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.FunctionReturnValueNotUsedInspectionResultFormat, Target.IdentifierName).Capitalize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
