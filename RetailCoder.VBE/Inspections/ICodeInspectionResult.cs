using System;

namespace Rubberduck.Inspections
{
    public interface ICodeInspectionResult
    {
        Rubberduck.VBA.Nodes.CommentNode Comment { get; }
        Antlr4.Runtime.ParserRuleContext Context { get; }
        Microsoft.Vbe.Interop.VBComponent FindComponent(Microsoft.Vbe.Interop.VBE vbe);
        System.Collections.Generic.IDictionary<string, Action<Microsoft.Vbe.Interop.VBE>> GetQuickFixes();
        string Name { get; }
        Rubberduck.Extensions.QualifiedSelection QualifiedSelection { get; }
        CodeInspectionSeverity Severity { get; }
    }
}
