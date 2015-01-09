using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspectionResult : CodeInspectionResultBase
    {
        public ImplicitVariantReturnTypeInspectionResult(string name, SyntaxTreeNode node, CodeInspectionSeverity severity)
            : base(name, node, severity)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                    {
                        {"Return explicit Variant", ReturnExplicitVariant}
                    }
                : new Dictionary<string, Action<VBE>>();
        }

        private void ReturnExplicitVariant(VBE vbe)
        {
            var instruction = Node.Instruction;
            if (!instruction.Line.IsMultiline)
            {
                var newContent = string.Concat(instruction.Value, " ", ReservedKeywords.As, " ", ReservedKeywords.Variant);
                var oldContent = instruction.Line.Content;

                var result = oldContent.Replace(instruction.Value, newContent);

                var module = vbe.FindCodeModules(instruction.Line.ProjectName, instruction.Line.ComponentName).First();
                module.ReplaceLine(instruction.Line.StartLineNumber, result);
                Handled = true;
            }
            else
            {
                // todo: implement for multiline
                throw new NotImplementedException("This method is not [yet] implemented for multiline instructions.");
            }
        }
    }
}