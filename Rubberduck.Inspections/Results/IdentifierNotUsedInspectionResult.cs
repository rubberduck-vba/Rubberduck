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
    public class IdentifierNotUsedInspectionResult : InspectionResultBase
    {
        public IdentifierNotUsedInspectionResult(IInspection inspection, Declaration target, ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context, target) {}

        public override string Description 
        {
            get
            {
                return string.Format(InspectionsUI.IdentifierNotUsedInspectionResultFormat, Target.DeclarationType.ToLocalizedString(), Target.IdentifierName).Capitalize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
