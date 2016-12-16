using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public abstract class SettingsViewModelBase : ViewModelBase
    {
        public CommandBase ExportButtonCommand { get; protected set; }

        public CommandBase ImportButtonCommand { get; protected set; }
    }
}
