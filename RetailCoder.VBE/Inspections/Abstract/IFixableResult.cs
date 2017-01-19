using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Abstract
{
    public interface IFixableResult : Parsing.Symbols.IInspectionResult, IComparable<IFixableResult>
    {
        IEnumerable<QuickFixBase> QuickFixes { get; }
        object[] ToArray();
    }
}
