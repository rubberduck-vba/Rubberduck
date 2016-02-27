using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public interface ICodeInspectionResult : IComparable<ICodeInspectionResult>, IComparable
    {
        IEnumerable<CodeInspectionQuickFix> QuickFixes { get; }
        CodeInspectionQuickFix DefaultQuickFix { get; }
        ParserRuleContext Context { get; }
        string Description { get; }
        QualifiedSelection QualifiedSelection { get; }
        IInspection Inspection { get; }
        string ToCsvString();
        object[] ToArray();
    }
}
