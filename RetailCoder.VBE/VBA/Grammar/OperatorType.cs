using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Grammar
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