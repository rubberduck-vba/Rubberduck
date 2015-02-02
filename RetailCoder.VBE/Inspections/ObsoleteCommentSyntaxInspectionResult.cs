using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using Rubberduck.Extensions;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCommentSyntaxInspectionResult(string inspection, CodeInspectionSeverity type, CommentNode comment) 
            : base(inspection, type, comment)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Replace Rem reserved keyword with single quote", ReplaceWithSingleQuote},
                    {"Remove comment", RemoveComment}
                };
        }

        private void ReplaceWithSingleQuote(VBE vbe)
        {
            throw new NotImplementedException();
        }

        private void RemoveComment(VBE vbe)
        {
            throw new NotImplementedException();
        }
    }
}