using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementUsageInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObsoleteCallStatementUsageInspectionResult(IInspection inspection, QualifiedContext<VBAParser.ExplicitCallStmtContext> qualifiedContext)
            : base(inspection, inspection.Description, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveExplicitCallStatemntQuickFix(Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class RemoveExplicitCallStatemntQuickFix : CodeInspectionQuickFix
    {
        public RemoveExplicitCallStatemntQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_RemoveObsoleteStatement)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;

            var selection = Context.GetSelection();
            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount);
            var originalInstruction = Context.GetText();

            var context = (VBAParser.ExplicitCallStmtContext)Context;

            string procedure;
            VBAParser.ArgsCallContext arguments;
            if (context.eCS_MemberProcedureCall() != null)
            {
                procedure = context.eCS_MemberProcedureCall().ambiguousIdentifier().GetText();
                arguments = context.eCS_MemberProcedureCall().argsCall();
            }
            else
            {
                procedure = context.eCS_ProcedureCall().ambiguousIdentifier().GetText();
                arguments = context.eCS_ProcedureCall().argsCall();
            }

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var argsList = (arguments == null
                ? new string[] { }
                : arguments.argCall().Select(e => e.GetText()))
            .ToList();
            var newInstruction = procedure + (argsList.Any() ? ' ' + string.Join(", ", argsList) : string.Empty);
            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
        }
    }
}