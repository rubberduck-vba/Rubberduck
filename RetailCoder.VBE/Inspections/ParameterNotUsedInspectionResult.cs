using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : CodeInspectionResultBase
    {
        public ParameterNotUsedInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, type, qualifiedName.ModuleScope, context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            // don't bother implementing this without implementing a ChangeSignatureRefactoring
            return new Dictionary<string, Action<VBE>>
            {
                //{"Remove unused parameter", RemoveUnusedParameter}
            };
        }
    }
}