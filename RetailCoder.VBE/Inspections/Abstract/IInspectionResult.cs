using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public interface IInspectionResult : IComparable<IInspectionResult>, IComparable
    {
        IEnumerable<QuickFixBase> QuickFixes { get; }
        string Description { get; }
        QualifiedSelection QualifiedSelection { get; }
        IInspection Inspection { get; }
        object[] ToArray();
        string ToClipboardString();
    }
}
