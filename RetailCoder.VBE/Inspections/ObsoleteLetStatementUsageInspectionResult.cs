using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementUsageInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteLetStatementUsageInspectionResult(string inspection, CodeInspectionSeverity type, 
            QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
        }

        private new VBAParser.LetStmtContext Context { get { return base.Context as VBAParser.LetStmtContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Remove obsolete statement", RemoveObsoleteStatement}
            };
        }

        private void RemoveObsoleteStatement(VBE vbe)
        {
            var module = QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var selection = Context.GetSelection();
            
            // remove line continuations to compare against Context:
            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                                          .Replace("\r\n", " ")
                                          .Replace("_", string.Empty);
            var originalInstruction = Context.GetText();

            var identifier = Context.implicitCallStmt_InStmt().GetText();
            var value = Context.valueStmt().GetText();

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = identifier + " = " + value;
            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
        }
    }
}