using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections
{
    /// <summary>
    /// An interface that abstracts a runnable code inspection.
    /// </summary>
    public interface IInspection : IInspectionModel, IComparable<IInspection>, IComparable
    {
        /// <summary>
        /// Runs code inspection on specified parse trees.
        /// </summary>
        /// <returns>Returns inspection results, if any.</returns>
        IEnumerable<InspectionResultBase> GetInspectionResults();

        /// <summary>
        /// Gets a string that contains additional/meta information about an inspection.
        /// </summary>
        string Meta { get; }

    }
}
