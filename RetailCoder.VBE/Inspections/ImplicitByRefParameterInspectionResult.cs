using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        public ImplicitByRefParameterInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<VBAParser.ArgContext> qualifiedContext)
            : base(inspection,type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
        }

        private new VBAParser.ArgContext Context { get { return base.Context as VBAParser.ArgContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            if ((Context.LPAREN() != null && Context.RPAREN() != null) || Context.PARAMARRAY() != null)
            {
                // array parameters & paramarrays must be passed by reference
                return new Dictionary<string, Action<VBE>>
                {
                    {"Pass parameter by reference explicitly", PassParameterByRef}
                };
            }

            return new Dictionary<string, Action<VBE>>
                {
                    {"Pass parameter by reference explicitly", PassParameterByRef},
                    {"Pass parameter by value", PassParameterByVal}
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

            var module = QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}