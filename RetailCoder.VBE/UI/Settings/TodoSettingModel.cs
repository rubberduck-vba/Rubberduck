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
        private BindingList<ToDoMarker> _markers;
        public BindingList<ToDoMarker> Markers { get { return _markers; } }

        public TodoSettingModel(List<ToDoMarker> markers)
        {
            _markers = new BindingList<ToDoMarker>(markers);
        }

        public void Save()
        {
            var settings = new ToDoListSettings(_markers.ToArray());
            var config = ConfigurationLoader.LoadConfiguration();
            config.UserSettings.ToDoListSettings = settings;

            ConfigurationLoader.SaveConfiguration<Configuration>(config);
        }
    }
}
