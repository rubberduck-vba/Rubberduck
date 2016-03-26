using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    public interface IHotkeyCommand
    {
        RubberduckHotkey Hotkey { get; }
    }
}