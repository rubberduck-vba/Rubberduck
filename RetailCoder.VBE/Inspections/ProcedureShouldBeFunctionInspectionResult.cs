using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ProcedureShouldBeFunctionInspectionResult : CodeInspectionResultBase
    {
       private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

       public ProcedureShouldBeFunctionInspectionResult(IInspection inspection, QualifiedContext<VBAParser.SubStmtContext> qualifiedContext)
            : base(inspection, string.Format(inspection.Description, qualifiedContext.Context.ambiguousIdentifier().GetText()), qualifiedContext.ModuleName, qualifiedContext.Context.ambiguousIdentifier())
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