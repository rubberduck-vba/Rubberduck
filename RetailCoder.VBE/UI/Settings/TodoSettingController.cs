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

        public TodoSettingController(ITodoSettingsView view, List<ToDoMarker> markers)
        {
            _markers = new BindingList<ToDoMarker>(markers);
            _view = view;

            _view.SelectionChanged += SelectionChanged;
            _view.TextChanged += TextChanged;
            _view.AddMarker += AddMarker;
            _view.RemoveMarker += RemoveMarker;
            _view.SaveMarker += SaveMarker;
        }

        private void SaveMarker(object sender, EventArgs e)
        {
            //todo: implement
            throw new NotImplementedException();

            //code behind implementation

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
            //_config.UserSettings.ToDoListSettings.ToDoMarkers = _markers.ToArray();
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _view.TodoMarkers.ToArray();
            _configService.SaveConfiguration<Configuration>(_config);
        }
    }
}
