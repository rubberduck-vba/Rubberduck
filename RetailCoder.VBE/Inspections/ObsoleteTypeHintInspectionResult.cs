using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteTypeHintInspectionResult(string inspection, CodeInspectionSeverity type,
            QualifiedContext<VBAParser.VariableSubStmtContext> qualifiedContext)
            : base(inspection, type, qualifiedContext.QualifiedName, qualifiedContext.Context)
        {
        }

        private new VBAParser.VariableSubStmtContext Context
        {
            get { return base.Context as VBAParser.VariableSubStmtContext; }
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                //{ "Remove type hint", null }
            };
        }
    }
}