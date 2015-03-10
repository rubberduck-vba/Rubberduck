using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        public ImplicitByRefParameterInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<VBAParser.ArgContext> qualifiedContext)
            : base(inspection,type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        private new VBAParser.ArgContext Context { get { return base.Context as VBAParser.ArgContext; } }

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
                    // this inspection doesn't know if parameter is assigned; suggest to pass ByRef explicitly
                    // and then let ParameterCanBeByVal inspection do its job.
                    //{"Pass parameter by value", PassParameterByVal},
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