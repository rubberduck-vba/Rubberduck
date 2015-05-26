using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteGlobalInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> context)
            : base(inspection, type, context.ModuleName, context.Context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {"Replace 'Global' access modifier with 'Public'", ChangeAccessModifier}
            };
        }

        private void ChangeAccessModifier()
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

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = Tokens.Public + ' ' + Context.GetText().Replace(Tokens.Global + ' ', string.Empty);
            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
        }
    }
}