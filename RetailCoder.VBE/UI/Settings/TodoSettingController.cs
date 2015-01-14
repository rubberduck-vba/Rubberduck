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

        public ToDoMarker ActiveMarker
        {
            get { return _view.TodoMarkers[_view.SelectedIndex]; }
        }

        public TodoSettingController(ITodoSettingsView view)
        {
            _view = view;

            if (_view.TodoMarkers != null)
            {
                _view.ActiveMarkerText = _view.TodoMarkers[0].Text;
                _view.ActiveMarkerPriority = _view.TodoMarkers[0].Priority;
            }

            _view.SelectionChanged += SelectionChanged;
            _view.TextChanged += TextChanged;
            _view.AddMarker += AddMarker;
            _view.RemoveMarker += RemoveMarker;
            _view.SaveMarker += SaveMarker;
            _view.PriorityChanged += PriorityChanged;
        }

        private void SaveMarker(object sender, EventArgs e)
        {
            //todo: add test
            var index = _view.SelectedIndex;
            _view.TodoMarkers[index].Text = _view.ActiveMarkerText;
            _view.TodoMarkers[index].Priority = _view.ActiveMarkerPriority;

            //old code behind implementation

            //var index = this.tokenListBox.SelectedIndex;
            //_markers[index].Text = tokenTextBox.Text;
            //_markers[index].Priority = (TodoPriority)priorityComboBox.SelectedIndex;
            //SaveActiveMarker(); //does this really need to happen? Changes still aren't being serialized.
        }

        private void RemoveMarker(object sender, EventArgs e)
        {
            _view.TodoMarkers.RemoveAt(_view.SelectedIndex);
        }

        private void AddMarker(object sender, EventArgs e)
        {
            var marker = new ToDoMarker(_view.ActiveMarkerText, _view.ActiveMarkerPriority);
            _view.TodoMarkers.Add(marker);

            _view.SelectedIndex = _view.TodoMarkers.Count - 1;
        }

        private void TextChanged(object sender, EventArgs e)
        {
            _view.SaveEnabled = true;
        }

        private void PriorityChanged(object sender, EventArgs e)
        {
            _view.SaveEnabled = true;
        }

        private void SelectionChanged(object sender, EventArgs e)
        {
            _view.ActiveMarkerPriority = this.ActiveMarker.Priority;
            _view.ActiveMarkerText = this.ActiveMarker.Text;

            _view.SaveEnabled = false;
        }

        public void SetActiveItem(int index)
        {
            _view.SelectedIndex = index;
        }

        [Obsolete]
        public void Save()
        {
            //_config.UserSettings.ToDoListSettings.ToDoMarkers = _markers.ToArray();
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _view.TodoMarkers.ToArray();
            _configService.SaveConfiguration<Configuration>(_config);
        }
    }
}
