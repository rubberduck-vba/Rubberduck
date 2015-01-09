using System.Runtime.InteropServices;

namespace Rubberduck.VBA
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