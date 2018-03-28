using System;
using System.Text;
using System.Windows.Forms;
using Rubberduck.UI;

namespace Rubberduck.Common.Hotkeys
{
    public struct HotkeyInfo
    {
        private const Keys Modifiers = Keys.Alt | Keys.Control | Keys.Shift;

        public HotkeyInfo(IntPtr hookId, Keys keys)
        {
            HookId = hookId;
            Keys = keys;
        }

        public IntPtr HookId { get; }
        public Keys Keys { get; }

        public override string ToString()
        {
            var builder = new StringBuilder();
            if (Keys.HasFlag(Keys.Alt))
            {
                builder.Append(RubberduckUI.GeneralSettings_HotkeyAlt);
                builder.Append('+');
            }
            if (Keys.HasFlag(Keys.Control))
            {
                builder.Append(RubberduckUI.GeneralSettings_HotkeyCtrl);
                builder.Append('+');
            }
            if (Keys.HasFlag(Keys.Shift))
            {
                builder.Append(RubberduckUI.GeneralSettings_HotkeyShift);
                builder.Append('+');
            }
            
            builder.Append(HotkeyDisplayConverter.Convert(Keys & ~Modifiers));
            return builder.ToString();
        }
    }
}
