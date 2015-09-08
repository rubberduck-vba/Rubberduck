using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementUsageInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObsoleteLetStatementUsageInspectionResult(string inspection, CodeInspectionSeverity type, 
            QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new[]
            {
                new RemoveExplicitLetStatementQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get {return _quickFixes; } }
    }

    public class RemoveExplicitLetStatementQuickFix : CodeInspectionQuickFix
    {
        public RemoveExplicitLetStatementQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_RemoveObsoleteStatement)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var selection = Context.GetSelection();
            var context = (VBAParser.LetStmtContext) Context;

            // remove line continuations to compare against context:
            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                                          .Replace("\r\n", " ")
                                          .Replace("_", string.Empty);
            var originalInstruction = Context.GetText();

            var identifier = context.implicitCallStmt_InStmt().GetText();
            var value = context.valueStmt().GetText();

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = identifier + " = " + value;
            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
        }
    }
}