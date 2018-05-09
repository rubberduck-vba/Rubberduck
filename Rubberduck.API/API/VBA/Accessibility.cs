using System.Runtime.InteropServices;

namespace Rubberduck.API.VBA
{
    [ComVisible(true), Guid(RubberduckGuid.AccessibilityGuid)]
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
