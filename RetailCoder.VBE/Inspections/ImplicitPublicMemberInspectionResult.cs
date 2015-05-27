using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspectionResult : CodeInspectionResultBase
    {
        public ImplicitPublicMemberInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection,type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {RubberduckUI.Inspections_SpecifyPublicModifierExplicitly,  SpecifyPublicModifier}
            };
        }

        private void SpecifyPublicModifier()
        {
            var oldContent = Context.GetText();
            var newContent = Tokens.Public + ' ' + oldContent;

            var selection = QualifiedSelection.Selection;

            var module = QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(oldContent, newContent);
            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, result);
        }
    }
}