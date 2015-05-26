using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public interface ICodeInspectionResult
    {
        IDictionary<string, Action> GetQuickFixes();
        ParserRuleContext Context { get; }
        string Name { get; }
        QualifiedSelection QualifiedSelection { get; }
        CodeInspectionSeverity Severity { get; }
    }
}
