using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
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