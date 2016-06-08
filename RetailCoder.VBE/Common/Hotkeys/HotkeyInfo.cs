using System;
using System.Text;
using System.Windows.Forms;

namespace Rubberduck.Common.Hotkeys
{
    public struct HotkeyInfo
    {
        private const Keys Modifiers = Keys.Alt | Keys.Control | Keys.Shift;

        private readonly IntPtr _hookId;
        private readonly Keys _keys;

        public HotkeyInfo(IntPtr hookId, Keys keys)
        {
            _hookId = hookId;
            _keys = keys;
        }

        public IntPtr HookId { get { return _hookId; } }
        public Keys Keys { get { return _keys; } }

        public override string ToString()
        {
            var builder = new StringBuilder();
            if (_keys.HasFlag(Keys.Alt))
            {
                builder.Append(Rubberduck.UI.RubberduckUI.GeneralSettings_HotkeyAlt);
                builder.Append('+');
            }
            if (_keys.HasFlag(Keys.Control))
            {
                builder.Append(Rubberduck.UI.RubberduckUI.GeneralSettings_HotkeyCtrl);
                builder.Append('+');
            }
            if (_keys.HasFlag(Keys.Shift))
            {
                builder.Append(Rubberduck.UI.RubberduckUI.GeneralSettings_HotkeyShift);
                builder.Append('+');
            }
            builder.Append(_keys & ~Modifiers);
            return builder.ToString();
        }
    }
}
