using System;
using System.Text;
using System.Windows.Forms;
using Rubberduck.Resources.Settings;

namespace Rubberduck.Common.Hotkeys
{
    public struct HotkeyInfo
    {
        private const Keys Modifiers = Keys.Alt | Keys.Control | Keys.Shift;

        public HotkeyInfo(ushort hookId, Keys keys)
        {
            HookId = hookId;
            Keys = keys;
        }

        public ushort HookId { get; }
        public Keys Keys { get; }

        public override string ToString()
        {
            var builder = new StringBuilder();
            if (Keys.HasFlag(Keys.Alt))
            {
                builder.Append(GeneralSettingsPage.HotkeyAlt);
                builder.Append('+');
            }
            if (Keys.HasFlag(Keys.Control))
            {
                builder.Append(GeneralSettingsPage.HotkeyCtrl);
                builder.Append('+');
            }
            if (Keys.HasFlag(Keys.Shift))
            {
                builder.Append(GeneralSettingsPage.HotkeyShift);
                builder.Append('+');
            }
            
            builder.Append(Rubberduck.UI.HotkeyDisplayConverter.Convert(Keys & ~Modifiers));
            return builder.ToString();
        }
    }
}
