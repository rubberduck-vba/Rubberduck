using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Config;
using System.ComponentModel;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingController
    {
        private Configuration _config;
        private IConfigurationService _configService;
        private ITodoSettingsView _view;

        private BindingList<ToDoMarker> _markers;
        public BindingList<ToDoMarker> Markers { get { return _markers; } }
        public ToDoMarker ActiveMarker
        {
            get { return _markers[_view.SelectedIndex]; }

        }

        [Obsolete]
        public TodoSettingController(IConfigurationService configService, ITodoSettingsView view)
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();
            _markers = new BindingList<ToDoMarker>(_config.UserSettings.ToDoListSettings.ToDoMarkers.ToList());
        }

        public TodoSettingController(ITodoSettingsView view, List<ToDoMarker> markers)
        {
            _markers = new BindingList<ToDoMarker>(markers);
            _view = view;

            _view.SelectionChanged += SelectionChanged;
        }

        private void SelectionChanged(object sender, EventArgs e)
        {
            _view.ActiveMarkerPriority = this.ActiveMarker.Priority;
            _view.ActiveMarkerText = this.ActiveMarker.Text;
        }


        public void SetActiveItem(int index)
        {
            _view.SelectedIndex = index;
        }

        public void Save()
        {
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _markers.ToArray();
            _configService.SaveConfiguration<Configuration>(_config);
        }
    }
}
