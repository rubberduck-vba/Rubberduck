using System;
using System.Collections.Generic;
using Rubberduck.Root;

namespace Rubberduck.Inspections.Abstract
{
    /// <summary>
    /// An interface that abstracts a runnable code inspection.
    /// </summary>
    public interface IInspection : IInspectionModel, IComparable<IInspection>, IComparable
    {
        /// <summary>
        /// Runs code inspection and returns inspection results.
        /// </summary>
        /// <returns>Returns inspection results, if any.</returns>
        [TimedCallIntercept]
        [EnumerableCounterIntercept]
        IEnumerable<InspectionResultBase> GetInspectionResults();

        /// <summary>
        /// Gets a string that contains additional/meta information about an inspection.
        /// </summary>
        string Meta { get; }
    }
}
