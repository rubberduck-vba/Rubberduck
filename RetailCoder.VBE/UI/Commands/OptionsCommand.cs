using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Commands
{
    public class OptionsCommand : ICommand
    {
        private readonly IGeneralConfigService _configService;

        public OptionsCommand(IGeneralConfigService configService)
        {
            _configService = configService;
        }

        public void Execute()
        {
            using (var window = new _SettingsDialog(_configService))
            {
                window.ShowDialog();
            }
        }
    }

    public class OptionsCommandMenuItem : CommandMenuItemBase
    {
        public OptionsCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_Options"; } }
    }
}