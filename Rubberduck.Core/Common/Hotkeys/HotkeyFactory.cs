using Rubberduck.Settings;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Common.Hotkeys
{
    public class HotkeyFactory
    {
        private readonly IEnumerable<CommandBase> _commands;

        public HotkeyFactory(IEnumerable<CommandBase> commands)
        {
            _commands = commands;
        }

        public Hotkey Create(HotkeySetting setting, IntPtr hWndVbe)
        {
            if (setting == null)
            {
                return null;
            }

            var commandToBind = _commands.FirstOrDefault(command => command.GetType().Name == setting.CommandTypeName);

            return commandToBind == null ? null : new Hotkey(hWndVbe, setting.ToString(), commandToBind);
        }
    }
}
