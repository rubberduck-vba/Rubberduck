using System.Runtime.InteropServices;

namespace Rubberduck.API.VBA
{
    [ComVisible(true)]
    public enum Accessibility
    {
        Implicit,
        Private,
        Public,
        Global,
        Friend,
        Static
    }
}
