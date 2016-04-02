using System.Windows.Forms;
using System.Windows.Input;

namespace Rubberduck.Common.Hotkeys
{
    public interface IHotkey : IAttachable
    {
        string Key { get; }
        ICommand Command { get; }
        HotkeyInfo HotkeyInfo { get; }
        Keys Combo { get; }
        Keys SecondKey { get; }
        bool IsTwoStepHotkey { get; }
    }
}