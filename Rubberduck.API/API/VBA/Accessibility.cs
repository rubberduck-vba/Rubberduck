using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

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
