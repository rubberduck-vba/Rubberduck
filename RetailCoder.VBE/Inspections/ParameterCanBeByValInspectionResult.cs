using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspectionResult : CodeInspectionResultBase
    {
        public ParameterCanBeByValInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, type, qualifiedName.ModuleScope, context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Pass parameter by value", PassParameterByVal}
            };
        }

        private void PassParameterByVal(VBE vbe)
        {
            ChangeParameterPassing(vbe, Tokens.ByVal);
        }

        private void ChangeParameterPassing(VBE vbe, string newValue)
        {
            var parameter = Context.GetText().Replace(Tokens.ByRef, string.Empty).Trim();
            var newContent = string.Concat(newValue, " ", parameter);
            var selection = QualifiedSelection.Selection;

            var module = vbe.FindCodeModules(QualifiedName.ProjectName, QualifiedName.ModuleName).First();
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}