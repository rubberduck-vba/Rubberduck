namespace Rubberduck.CodeAnalysis.Inspections
{
    public enum CodeInspectionSeverity
    {
        /// <summary>
        /// Inspection will not run.
        /// </summary>
        DoNotShow,
        /// <summary>
        /// Low severity setting.
        /// </summary>
        Hint,
        /// <summary>
        /// Medium-low severity setting.
        /// </summary>
        Suggestion,
        /// <summary>
        /// Medium-high severity setting.
        /// </summary>
        Warning,
        /// <summary>
        /// High severity setting.
        /// </summary>
        Error
    }
}
