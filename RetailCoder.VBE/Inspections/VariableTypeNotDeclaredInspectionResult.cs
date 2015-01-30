using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class VariableTypeNotDeclaredInspectionResult : CodeInspectionResultBase
    {
        public VariableTypeNotDeclaredInspectionResult(string inspection, ParserRuleContext context, CodeInspectionSeverity type, string project, string module)
            : base(inspection, context, type, project, module)
        {
            _project = project;
            _module = module;
        }

        private readonly string _project;
        private readonly string _module;

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
            var name = (string)((dynamic)Context).ambiguousIdentifier().GetText(); // casts dynamic context away
            var newContent = name + " " + ReservedKeywords.As + " " + ReservedKeywords.Variant;
            var oldContent = Context.GetLine();

            var result = oldContent.Replace(name, newContent);
            var module = vbe.FindCodeModules(_project, _module).First();
            module.ReplaceLine(Context.Start.Line, result);
        }
    }
}