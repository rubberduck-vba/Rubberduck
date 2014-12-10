using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspectionResult : CodeInspectionResultBase
    {
        public ImplicitVariantReturnTypeInspectionResult(string name, Instruction instruction, CodeInspectionSeverity severity)
            : base(name, instruction, severity)
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
            if (!Instruction.Line.IsMultiline)
            {
                var newContent = string.Concat(Instruction.Value, " ", ReservedKeywords.As, " ", ReservedKeywords.Variant);
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
    }
}