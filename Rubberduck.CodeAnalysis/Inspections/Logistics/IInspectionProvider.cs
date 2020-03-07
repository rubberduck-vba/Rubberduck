using System.Collections.Generic;

namespace Rubberduck.CodeAnalysis.Inspections.Logistics
{
    internal interface IInspectionProvider
    {
        IEnumerable<IInspection> Inspections { get; }
    }
}