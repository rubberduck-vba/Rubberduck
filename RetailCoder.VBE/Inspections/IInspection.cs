using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    /// <summary>
    /// An interface that abstracts a code inspection.
    /// </summary>
    [ComVisible(false)]
    public interface IInspection
    {
        /// <summary>
        /// Gets a short description for the code inspection.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets a short message that describes how a code issue can be fixed.
        /// </summary>
        string QuickFixMessage { get; }

        /// <summary>
        /// Gets a value indicating the type of the code inspection.
        /// </summary>
        CodeInspectionType InspectionType { get; }

        /// <summary>
        /// Gets a value indicating the severity level of the code inspection.
        /// </summary>
        CodeInspectionSeverity Severity { get; set; }

        /// <summary>
        /// Gets/sets a valud indicating whether the inspection is enabled or not.
        /// </summary>
        bool IsEnabled { get; set; }

        /// <summary>
        /// Runs code inspection on specified tree node (and child nodes).
        /// </summary>
        /// <param name="node">The <see cref="SyntaxTreeNode"/> to analyze.</param>
        /// <returns>Returns inspection results, if any.</returns>
        IEnumerable<CodeInspectionResultBase> Inspect(SyntaxTreeNode node);
    }
}
