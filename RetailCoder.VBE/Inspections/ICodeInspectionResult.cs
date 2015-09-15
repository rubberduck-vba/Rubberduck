using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public interface ICodeInspectionResult
    {
        IEnumerable<CodeInspectionQuickFix> QuickFixes { get; }
        ParserRuleContext Context { get; }
        string Name { get; }
        QualifiedSelection QualifiedSelection { get; }
        IInspection Inspection { get; }
    }
}
