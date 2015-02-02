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
        public VariableTypeNotDeclaredInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
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
            throw new NotImplementedException();
        }
    }
}