using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class EmptyStringLiteralInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public EmptyStringLiteralInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, inspection.Description, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new[]
            {
                new RepaceEmptyStringLiteralStatementQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class RepaceEmptyStringLiteralStatementQuickFix : CodeInspectionQuickFix
    {
        public RepaceEmptyStringLiteralStatementQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.EmptyStringLiteralInspectionQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var literal = (VBAParser.LiteralContext)Context;
            var newCodeLines = module.Lines[literal.Start.Line, 1].Replace("\"\"", "vbNullString");

            module.ReplaceLine(literal.Start.Line, newCodeLines);
        }
    }
}