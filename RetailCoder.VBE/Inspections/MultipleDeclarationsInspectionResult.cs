using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class MultipleDeclarationsInspectionResult : CodeInspectionResultBase
    {
        public MultipleDeclarationsInspectionResult(string inspection, CodeInspectionSeverity type, 
            QualifiedContext<VisualBasic6Parser.VariableListStmtContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
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
            foreach (var node in Context.children)
            {
                newContent.AppendLine(indent + node.GetText());
            }

            var module = vbe.FindCodeModules(QualifiedName).First();
            module.ReplaceLine(Context.GetSelection().StartLine, newContent.ToString());
        }
    }
}