using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public interface ICodeInspectionResult
    {
        CommentNode Comment { get; }
        ParserRuleContext Context { get; }
        VBComponent FindComponent(VBE vbe);
        IDictionary<string, Action<VBE>> GetQuickFixes();
        string Name { get; }
        QualifiedSelection QualifiedSelection { get; }
        CodeInspectionSeverity Severity { get; }
    }
}
