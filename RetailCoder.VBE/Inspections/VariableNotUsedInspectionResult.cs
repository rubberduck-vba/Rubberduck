using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Inspections
{
    public class VariableNotUsedInspectionResult : CodeInspectionResultBase
    {
        public VariableNotUsedInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, type, qualifiedName, context)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Remove unused variable declaration", RemoveUnusedDeclaration}
                };
        }

        private void RemoveUnusedDeclaration(VBE vbe)
        {
            
        }
    }
}