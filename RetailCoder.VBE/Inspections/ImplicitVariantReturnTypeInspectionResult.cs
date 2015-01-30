using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspectionResult : CodeInspectionResultBase
    {
        public ImplicitVariantReturnTypeInspectionResult(string name, ParserRuleContext context, CodeInspectionSeverity severity, string project, string module, string procedure)
            : base(name, context, severity, project, module)
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
                    {"Return explicit Variant", ReturnExplicitVariant}
                };
        }

        private void ReturnExplicitVariant(VBE vbe)
        {
            var instruction = Context.GetLine();
            var newContent = instruction + " " + ReservedKeywords.As + " " + ReservedKeywords.Variant;
            var oldContent = instruction;

            var result = oldContent.Replace(instruction, newContent);

            var module = vbe.FindCodeModules(_project, _module).First();
            module.ReplaceLine(Context.GetSelection().StartLine, result);
        }
    }
}