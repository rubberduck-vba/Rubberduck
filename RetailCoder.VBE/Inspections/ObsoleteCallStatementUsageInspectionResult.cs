using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementUsageInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCallStatementUsageInspectionResult(string inspection, CodeInspectionSeverity type,
            QualifiedContext<VBAParser.ExplicitCallStmtContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
        }

        private new VBAParser.ExplicitCallStmtContext Context { get { return base.Context as VBAParser.ExplicitCallStmtContext;} }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {"Remove obsolete statement", RemoveObsoleteStatement}
            };
        }

        private void RemoveObsoleteStatement()
        {
            var module = QualifiedName.Component.CodeModule;

            var selection = Context.GetSelection();
            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount);
            var originalInstruction = Context.GetText();

            string procedure;
            VBAParser.ArgsCallContext arguments;
            if (Context.eCS_MemberProcedureCall() != null)
            {
                procedure = Context.eCS_MemberProcedureCall().ambiguousIdentifier().GetText();
                arguments = Context.eCS_MemberProcedureCall().argsCall();
            }
            else
            {
                procedure = Context.eCS_ProcedureCall().ambiguousIdentifier().GetText();
                arguments = Context.eCS_ProcedureCall().argsCall();
            }

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var argsList = arguments == null
                ? new[] { string.Empty }
                : arguments.argCall().Select(e => e.GetText());
            var newInstruction = procedure + ' ' + string.Join(", ", argsList);
            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
        }
    }
}