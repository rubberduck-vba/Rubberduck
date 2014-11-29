using System.Runtime.InteropServices;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public enum CodeInspectionType
    {
        MaintainabilityAndReadabilityIssues,
        CodeQualityIssues
    }
}