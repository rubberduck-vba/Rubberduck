using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IInspectionResult : IComparable<IInspectionResult>, IComparable
    {
        IEnumerable<IQuickFix> QuickFixes { get; }
        string Description { get; }
        QualifiedSelection QualifiedSelection { get; }
        IInspection Inspection { get; }
        Declaration Target { get; }
    }
}
