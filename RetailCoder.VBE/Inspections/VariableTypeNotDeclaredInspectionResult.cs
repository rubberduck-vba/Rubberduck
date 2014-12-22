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
    public class VariableTypeNotDeclaredInspectionResult : CodeInspectionResultBase
    {
        private readonly IdentifierNode _identifier;

        public VariableTypeNotDeclaredInspectionResult(string inspection, IdentifierNode identifier,
            CodeInspectionSeverity type)
            : this(inspection, identifier.Instruction, type)
        {
            _identifier = identifier;
        }

        private VariableTypeNotDeclaredInspectionResult(string inspection, Instruction instruction,
            CodeInspectionSeverity type)
            : base(inspection, instruction, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                    {
                        {"Declare as explicit Variant", DeclareAsExplicitVariant}
                    }
                : new Dictionary<string, Action<VBE>>();
        }

        private void DeclareAsExplicitVariant(VBE vbe)
        {
            var newContent = string.Concat(_identifier.Name, " ", ReservedKeywords.As, " ", ReservedKeywords.Variant);
            var oldContent = Instruction.Line.Content;

            var result = oldContent.Replace(_identifier.Name, newContent);
            var module = vbe.FindCodeModules(Instruction.Line.ProjectName, Instruction.Line.ComponentName).First();
            module.ReplaceLine(Instruction.Line.StartLineNumber, result);
            Handled = true;
        }
    }
}