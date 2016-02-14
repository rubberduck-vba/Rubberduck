using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MultilineParameterInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public MultilineParameterInspectionResult(IInspection inspection, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = new[]
            {
                new MakeSingleLineParameterQuickFix(Context, QualifiedSelection), 
            };
        }

        public override string Description
        {
            get
            {
                return string.Format(
                    Target.Context.GetSelection().LineCount > 3
                        ? RubberduckUI.EasterEgg_Continuator
                        : InspectionsUI.MultilineParameterInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }
}