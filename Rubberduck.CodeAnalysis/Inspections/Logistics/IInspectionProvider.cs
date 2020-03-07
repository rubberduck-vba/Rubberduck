using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.Inspections
{
    public interface IInspectionProvider
    {
        IEnumerable<IInspection> Inspections { get; }
    }
}