using System.Windows.Forms;

namespace Rubberduck.Common.Hotkeys
{
    public interface IHotkey : IAttachable
    {
        string Key { get; }
        HotkeyInfo HotkeyInfo { get; }
        Keys Combo { get; }
        Keys SecondKey { get; }
        bool IsTwoStepHotkey { get; }
    }
}