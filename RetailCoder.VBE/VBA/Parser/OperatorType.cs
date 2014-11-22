using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public enum OperatorType
    {
        Assignment,
        Comparison,
        Logical,
        Arithmetic,
        Concatenation
    }
}