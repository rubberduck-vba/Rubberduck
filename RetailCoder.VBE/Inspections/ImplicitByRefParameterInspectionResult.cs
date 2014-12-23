using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        public ImplicitByRefParameterInspectionResult(string inspection, SyntaxTreeNode node, CodeInspectionSeverity type)
            : base(inspection, node, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            if ((Node as ParameterNode).Identifier.IsArray)
            {
                // array parameters must be passed by reference
                return new Dictionary<string, Action<VBE>>
                {
                    {"Pass parameter by reference explicitly", PassParameterByRef}
                };
            }

            return new Dictionary<string, Action<VBE>>
                {
                    {"Pass parameter by value", PassParameterByVal},
                    {"Pass parameter by reference explicitly", PassParameterByRef}
                };
        }

        private void PassParameterByRef(VBE vbe)
        {
            ChangeParameterPassing(vbe, ReservedKeywords.ByRef);
        }

        private void PassParameterByVal(VBE vbe)
        {
            ChangeParameterPassing(vbe, ReservedKeywords.ByVal);
        }

        private void ChangeParameterPassing(VBE vbe, string newValue)
        {
            var instruction = Node.Instruction;
            if (!instruction.Line.IsMultiline)
            {
                var newContent = string.Concat(newValue, " ", instruction.Value);
                var oldContent = instruction.Line.Content;

                var result = oldContent.Replace(instruction.Value, newContent);

                var module = vbe.FindCodeModules(instruction.Line.ProjectName, instruction.Line.ComponentName).First();
                module.ReplaceLine(instruction.Line.StartLineNumber, result);
                Handled = true;
            }
            else
            {
                // todo: implement for multiline
                throw new NotImplementedException("This method is not yet implemented for multiline instructions.");
            }
        }
    }
}