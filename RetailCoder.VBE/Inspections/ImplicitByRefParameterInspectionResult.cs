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
        public ImplicitByRefParameterInspectionResult(string inspection, Instruction instruction, CodeInspectionSeverity type)
            : base(inspection, instruction, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                    {
                        {"Pass parameter by value", PassParameterByVal},
                        {"Pass parameter by reference explicitly", PassParameterByRef}
                    }
                : new Dictionary<string, Action<VBE>>();
        }

        private void PassParameterByRef(VBE vbe)
        {
            if (!Instruction.Line.IsMultiline)
            {
                var newContent = string.Concat(ReservedKeywords.ByRef, " ", Instruction.Value);
                var oldContent = Instruction.Line.Content;

                var result = oldContent.Replace(Instruction.Value, newContent);

                var module = vbe.FindCodeModules(Instruction.Line.ProjectName, Instruction.Line.ComponentName).First();
                module.ReplaceLine(Instruction.Line.StartLineNumber, result);
                Handled = true;
            }
            else
            {
                // todo: implement for multiline
                throw new NotImplementedException("This method is not [yet] implemented for multiline instructions.");
            }
        }

        private void PassParameterByVal(VBE vbe)
        {
            if (!Instruction.Line.IsMultiline)
            {
                var newContent = string.Concat(ReservedKeywords.ByVal, " ", Instruction.Value);
                var oldContent = Instruction.Line.Content;

                var result = oldContent.Replace(Instruction.Value, newContent);

                var module = vbe.FindCodeModules(Instruction.Line.ProjectName, Instruction.Line.ComponentName).First();
                module.ReplaceLine(Instruction.Line.StartLineNumber, result);
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