namespace Rubberduck.Inspections
{
    /// <summary>
    /// An interface that abstracts the data structure for a code inspection
    /// </summary>
    public interface IInspectionModel
    {
        /// <summary>
        /// Gets a short description for the code inspection.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets a value indicating the type of the code inspection.
        /// </summary>
        CodeInspectionType InspectionType { get; }

        /// <summary>
        /// Gets a value indicating the severity level of the code inspection.
        /// </summary>
        CodeInspectionSeverity Severity { get; set; }
    }
}