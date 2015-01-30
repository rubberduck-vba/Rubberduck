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
        public ImplicitByRefParameterInspectionResult(string inspection, VisualBasic6Parser.ArgContext context, CodeInspectionSeverity type,string project, string module, string procedure)
            : base(inspection, context, type, project, module)
        {
            _project = project;
            _module = module;
        }

        private readonly string _project;
        private readonly string _module;

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
            ChangeParameterPassing(vbe, ReservedKeywords.ByRef);
        }

        private void PassParameterByVal(VBE vbe)
        {
            ChangeParameterPassing(vbe, ReservedKeywords.ByVal);
        }

        private void ChangeParameterPassing(VBE vbe, string newValue)
        {
            var oldContent = Context.GetText();
            var newContent = string.Concat(newValue, " ", Context.GetText());

            var result = oldContent.Replace(oldContent, newContent);

            var module = vbe.FindCodeModules(_project, _module).First();
            module.ReplaceLine(Context.GetSelection().StartLine, result);
        }
    }
}