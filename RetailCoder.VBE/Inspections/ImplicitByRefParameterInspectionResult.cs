using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        public ImplicitByRefParameterInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<VisualBasic6Parser.ArgContext> qualifiedContext)
            : base(inspection,type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        private new VisualBasic6Parser.ArgContext Context { get { return base.Context as VisualBasic6Parser.ArgContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            if (Context.LPAREN() != null && Context.RPAREN() != null)
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
            ChangeParameterPassing(vbe, Tokens.ByRef);
        }

        private void PassParameterByVal(VBE vbe)
        {
            ChangeParameterPassing(vbe, Tokens.ByVal);
        }

        private void ChangeParameterPassing(VBE vbe, string newValue)
        {
            var parameter = Context.GetText();
            var newContent = string.Concat(newValue, " ", parameter);
            var selection = QualifiedSelection.Selection;

            var module = vbe.FindCodeModules(QualifiedName.ProjectName, QualifiedName.ModuleName).First();
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}