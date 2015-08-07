using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Command
{
    public class OptionsCommand : ICommand
    {
        public void Execute()
        {
            using (var window = new SettingsDialog())
            {
                window.ShowDialog();
            }
        }
    }

    public class OptionsMenuCommand : CommandMenuItemBase
    {
        public OptionsMenuCommand(ICommand command) : base(command)
        {
        }

        public override string Key
        {
            get { return RubberduckUI.RubberduckMenu_Options; }
        }

        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Options; } }
    }
}
