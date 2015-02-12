using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementUsageInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCallStatementUsageInspectionResult(string inspection, CodeInspectionSeverity type,
            QualifiedContext<VisualBasic6Parser.ExplicitCallStmtContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        private new VisualBasic6Parser.ExplicitCallStmtContext Context { get { return base.Context as VisualBasic6Parser.ExplicitCallStmtContext;} }

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