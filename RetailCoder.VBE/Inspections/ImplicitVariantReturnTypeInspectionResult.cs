using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspectionResult : CodeInspectionResultBase
    {
        public ImplicitVariantReturnTypeInspectionResult(string name, CodeInspectionSeverity severity, 
            QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(name, severity, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

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

            var module = vbe.FindCodeModules(QualifiedName).First();
            module.ReplaceLine(Context.GetSelection().StartLine, result);
        }
    }
}