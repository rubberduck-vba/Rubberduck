using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class VariableTypeNotDeclaredInspectionResult : CodeInspectionResultBase
    {
        public VariableTypeNotDeclaredInspectionResult(string inspection, CodeInspectionSeverity type, ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, type, qualifiedName, context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Declare as explicit Variant", DeclareAsExplicitVariant}
                };
        }

        private void DeclareAsExplicitVariant(VBE vbe)
        {
            var component = FindComponent(vbe);
            if (component == null)
            {
                throw new InvalidOperationException("'" + QualifiedName + "' not found.");
            }

            var codeLine = component.CodeModule.get_Lines(QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.LineCount);
            var instruction = Context.GetText();
            var fixedCodeLine = codeLine.Replace(instruction, instruction + " " + Tokens.As + " " + Tokens.Variant);

            component.CodeModule.ReplaceLine(QualifiedSelection.Selection.StartLine, fixedCodeLine);
        }
    }
}