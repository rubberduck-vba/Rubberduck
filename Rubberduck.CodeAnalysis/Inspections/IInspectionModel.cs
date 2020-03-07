namespace Rubberduck.CodeAnalysis.Inspections
{
    /// <summary>
    /// An interface that abstracts the data structure for a code inspection
    /// </summary>
    public interface IInspectionModel
    {
        /// <summary>
        /// Gets the inspection name.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets a string referring to this inspection in VBA @annotations.
        /// </summary>
        string AnnotationName { get; }

        /// <summary>
        /// Gets a short description for the code inspection.
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Gets a value indicating the type of the code inspection.
        /// </summary>
        CodeInspectionType InspectionType { get; set; }

        /// <summary>
        /// Gets a value indicating the severity level of the code inspection.
        /// </summary>
        CodeInspectionSeverity Severity { get; set; }

        /// <summary>
        /// Gets a string that contains additional/meta information about an inspection.
        /// </summary>
        // ReSharper disable once UnusedMember.Global; property is used in XAML bindings.
        string Meta { get; }
    }
}
