using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Abstract
{
    public interface IInspectionResult : Parsing.Symbols.IInspectionResult, IComparable<IInspectionResult>
    {
        IEnumerable<QuickFixBase> QuickFixes { get; }
        object[] ToArray();
    }
}
