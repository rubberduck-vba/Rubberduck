using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class MultipleDeclarationsInspectionResult : CodeInspectionResultBase
    {
        public MultipleDeclarationsInspectionResult(string inspection, SyntaxTreeNode node, CodeInspectionSeverity type)
            : base(inspection, node, type)
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
            var instruction = Node.Instruction;
            var newContent = new StringBuilder();
            var indent = new string(' ', instruction.StartColumn - 1);
            foreach (var node in Node.ChildNodes.Cast<IdentifierNode>())
            {
                newContent.AppendLine(indent + node);
            }

            var module = vbe.FindCodeModules(instruction.Line.ProjectName, instruction.Line.ComponentName).First();
            module.ReplaceLine(instruction.Line.StartLineNumber, newContent.ToString());
        }
    }
}