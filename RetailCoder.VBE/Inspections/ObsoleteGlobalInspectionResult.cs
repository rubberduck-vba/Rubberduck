using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteGlobalInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> context)
            : base(inspection, type, context.QualifiedName, context.Context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Replace 'Global' access modifier with 'Public'", ChangeAccessModifier}
            };
        }

        private void ChangeAccessModifier(VBE vbe)
        {
            var module = vbe.FindCodeModules(QualifiedName).SingleOrDefault();
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