using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementUsageInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteLetStatementUsageInspectionResult(string inspection, CodeInspectionSeverity type, 
            QualifiedContext<VisualBasic6Parser.LetStmtContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        private new VisualBasic6Parser.LetStmtContext Context { get { return base.Context as VisualBasic6Parser.LetStmtContext; } }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                {"Remove redundant keyword", RemoveRedundantKeyword}
            };
        }

        private void RemoveRedundantKeyword(VBE vbe)
        {
            throw new NotImplementedException();
        }
    }
}