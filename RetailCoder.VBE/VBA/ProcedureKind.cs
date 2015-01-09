using System.Runtime.InteropServices;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public enum ProcedureKind
    {
        Sub,
        Function,
        PropertyGet,
        PropertyLet,
        PropertySet
    }
}