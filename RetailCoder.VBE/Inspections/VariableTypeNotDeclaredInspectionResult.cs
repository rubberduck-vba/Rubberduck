using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class VariableTypeNotDeclaredInspectionResult : CodeInspectionResultBase
    {
        private readonly IdentifierNode _identifier;

        public VariableTypeNotDeclaredInspectionResult(string inspection, IdentifierNode identifier,
            CodeInspectionSeverity type)
            : this(inspection, identifier as SyntaxTreeNode, type)
        {
            _identifier = identifier;
        }

        private VariableTypeNotDeclaredInspectionResult(string inspection, SyntaxTreeNode node,
            CodeInspectionSeverity type)
            : base(inspection, node, type)
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
            var instruction = Node.Instruction;
            var newContent = string.Concat(_identifier.Name, " ", ReservedKeywords.As, " ", ReservedKeywords.Variant);
            var oldContent = instruction.Line.Content;

            var result = oldContent.Replace(_identifier.Name, newContent);
            var module = vbe.FindCodeModules(instruction.Line.ProjectName, instruction.Line.ComponentName).First();
            module.ReplaceLine(instruction.Line.StartLineNumber, result);
            Handled = true;
        }
    }
}