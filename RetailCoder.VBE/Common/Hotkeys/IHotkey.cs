using System.Windows.Forms;
using Rubberduck.UI.Command;

namespace Rubberduck.Common.Hotkeys
{
    public interface IHotkey : IAttachable
    {
        string Key { get; }
        CommandBase Command { get; }
        HotkeyInfo HotkeyInfo { get; }
        Keys Combo { get; }
        Keys SecondKey { get; }
        bool IsTwoStepHotkey { get; }
    }
}
