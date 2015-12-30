using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ProcedureShouldBeFunctionInspectionResult : CodeInspectionResultBase
    {
       private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

       public ProcedureShouldBeFunctionInspectionResult(IInspection inspection, string result, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, result, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = new[]
            {
                new ChangeProcedureToFunction(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class ChangeProcedureToFunction : CodeInspectionQuickFix
    {
        public ChangeProcedureToFunction(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.ProcedureShouldBeFunctionInspectionQuickFix)
        {
        }

        public override void Fix()
        {
        }
    }
}