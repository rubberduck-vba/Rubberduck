using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Command
{
    public class OptionsCommand : ICommand
    {
        private readonly IGeneralConfigService _service;
        public OptionsCommand(IGeneralConfigService service)
        {
            _service = service;
        }

        public void Execute()
        {
            using (var window = new SettingsDialog(_service))
            {
                window.ShowDialog();
            }
        }
    }
}
