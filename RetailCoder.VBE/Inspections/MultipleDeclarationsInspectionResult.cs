using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class MultipleDeclarationsInspectionResult : CodeInspectionResultBase
    {
        public MultipleDeclarationsInspectionResult(string inspection, ParserRuleContext context, CodeInspectionSeverity type, string project, string module)
            : base(inspection, context, type, project, module)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Separate multiple declarations into multiple instructions", SplitDeclarations},
            };
        }

        private void SplitDeclarations(VBE vbe)
        {
            var newContent = new StringBuilder();
            var indent = new string(' ', Context.GetSelection().StartColumn);
            foreach (var node in Context.GetDeclarations())
            {
                var name = (string)((dynamic)Context).ambiguousIdentifier().GetText();
                newContent.AppendLine(indent + node);
            }

            var module = vbe.FindCodeModules(ProjectName, ModuleName).First();
            module.ReplaceLine(Context.GetSelection().StartLine, newContent.ToString());
        }
    }
}