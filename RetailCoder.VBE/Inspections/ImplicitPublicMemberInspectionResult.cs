using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitPublicMemberInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection,type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new[]
            {
                new SpecifyExplicitPublicModifierQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class SpecifyExplicitPublicModifierQuickFix : CodeInspectionQuickFix
    {
        public SpecifyExplicitPublicModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_SpecifyPublicModifierExplicitly)
        {
        }

        public override void Fix()
        {
            var oldContent = Context.GetText();
            var newContent = Tokens.Public + ' ' + oldContent;

            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(oldContent, newContent);
            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}