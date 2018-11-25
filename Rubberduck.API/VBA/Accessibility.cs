using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;
using Source = Rubberduck.Parsing.Symbols;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.AccessibilityGuid)
    ]
    public enum Accessibility
    {
        Private = Source.Accessibility.Private,
        Friend = Source.Accessibility.Friend,
        Global = Source.Accessibility.Global,
        Implicit = Source.Accessibility.Implicit,
        Public = Source.Accessibility.Public,
        Static = Source.Accessibility.Static
    }
}
