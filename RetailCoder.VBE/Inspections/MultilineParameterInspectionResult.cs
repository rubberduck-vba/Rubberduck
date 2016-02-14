using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MultilineParameterInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public MultilineParameterInspectionResult(IInspection inspection, string result, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, result, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = new[]
            {
                new MakeSingleLineParameterQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class MakeSingleLineParameterQuickFix : CodeInspectionQuickFix
    {
        public MakeSingleLineParameterQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_MultilineParameter)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Selection.Selection;

            var lines = module.Lines[selection.StartLine, selection.EndLine - selection.StartLine + 1];

            var startLine = module.Lines[selection.StartLine, 1];
            var endLine = module.Lines[selection.EndLine, 1];

            var adjustedStartColumn = selection.StartColumn - 1;
            var adjustedEndColumn = lines.Length - (endLine.Length - (selection.EndColumn > endLine.Length ? endLine.Length : selection.EndColumn - 1));

            var parameter = lines.Substring(adjustedStartColumn,
                adjustedEndColumn - adjustedStartColumn)
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            var start = startLine.Remove(adjustedStartColumn);
            var end = lines.Remove(0, adjustedEndColumn);

            module.ReplaceLine(selection.StartLine, start + parameter + end);
            module.DeleteLines(selection.StartLine + 1, selection.EndLine - selection.StartLine);
        }
    }
}