using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class RemoveExplicitCallStatmentQuickFix : CodeInspectionQuickFix
    {
        public RemoveExplicitCallStatmentQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveObsoleteStatementQuickFix)
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