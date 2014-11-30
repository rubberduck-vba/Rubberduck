using System.Runtime.InteropServices;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public enum CodeInspectionSeverity
    {
        DoNotShow,
        Hint,
        Suggestion,
        Warning,
        Error
    }
}