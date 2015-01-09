using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Grammar
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