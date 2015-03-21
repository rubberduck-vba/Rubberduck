using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;

namespace Rubberduck.Inspections
{
    public class ModuleOptionsNotSpecifiedFirstInspectionResult : CodeInspectionResultBase
    {
        public ModuleOptionsNotSpecifiedFirstInspectionResult(string inspection, CodeInspectionSeverity type, CommentNode comment) 
            : base(inspection, type, comment)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                // todo: implement quickfix?
            };
        }
    }
}