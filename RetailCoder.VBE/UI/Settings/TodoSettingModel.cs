using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Config;
using System.ComponentModel;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingModel
    {
        private Configuration _config;
        private IConfigurationService _configService;

        private BindingList<ToDoMarker> _markers;
        public BindingList<ToDoMarker> Markers { get { return _markers; } }

        public TodoSettingModel(IConfigurationService configService)
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();
            _markers = new BindingList<ToDoMarker>(_config.UserSettings.ToDoListSettings.ToDoMarkers.ToList());
        }

        public void Save()
        {
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _markers.ToArray();
            _configService.SaveConfiguration<Configuration>(_config);
        }
    }
}
