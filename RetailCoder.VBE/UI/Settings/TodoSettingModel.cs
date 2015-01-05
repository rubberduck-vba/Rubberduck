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

        private BindingList<ToDoMarker> _markers;
        public BindingList<ToDoMarker> Markers { get { return _markers; } }

        public TodoSettingModel(Configuration config)
        {
            _config = config; //todo: set config from IConfigurationService
            _markers = new BindingList<ToDoMarker>(config.UserSettings.ToDoListSettings.ToDoMarkers.ToList());
        }

        public void Save()
        {
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _markers.ToArray();
            ConfigurationLoader.SaveConfiguration<Configuration>(_config);
            //todo: create IConfigurationService and inject one in lieu of ConfigurationLoader
        }
    }
}
