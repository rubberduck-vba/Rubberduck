using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public abstract class SettingsViewModelBase : ViewModelBase
    {
        public CommandBase ExportButtonCommand { get; protected set; }

        public CommandBase ImportButtonCommand { get; protected set; }
    }
}
