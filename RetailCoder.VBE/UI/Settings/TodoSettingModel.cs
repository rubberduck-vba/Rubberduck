using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingModel
    {
        private List<ToDoMarker> _markers;
        public List<ToDoMarker> Markers { get { return _markers; } }

        public TodoSettingModel(List<ToDoMarker> markers)
        {
            _markers = markers;
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
