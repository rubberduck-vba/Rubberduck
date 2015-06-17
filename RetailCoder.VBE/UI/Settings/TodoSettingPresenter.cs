using System;
using System.ComponentModel;
using System.Linq;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingPresenter
    {
        private readonly ITodoSettingsView _view;

        public ToDoMarker ActiveMarker
        {
            get { return _view.TodoMarkers[_view.SelectedIndex]; }
        }

        public TodoSettingPresenter(ITodoSettingsView view)
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
            //todo: add test; How? I can't click the save button. Code smell here.
            var index = _view.SelectedIndex;
            _view.TodoMarkers[index].Text = _view.ActiveMarkerText;
            _view.TodoMarkers[index].Priority = _view.ActiveMarkerPriority;
        }

        private void RemoveMarker(object sender, EventArgs e)
        {
            var oldList = _view.TodoMarkers.ToList();
            oldList.RemoveAt(_view.SelectedIndex);
            _view.TodoMarkers = new BindingList<ToDoMarker>(oldList);
        }

        private void AddMarker(object sender, EventArgs e)
        {
            var oldList = _view.TodoMarkers.ToList();
            var marker = new ToDoMarker(_view.ActiveMarkerText, _view.ActiveMarkerPriority);
            oldList.Add(marker);

            _view.TodoMarkers = new BindingList<ToDoMarker>(oldList);

            _view.SelectedIndex = _view.TodoMarkers.Count - 1;
        }

        private void TextChanged(object sender, EventArgs e)
        {
            var markerExists = _view.TodoMarkers.Where((t, i) => i != _view.SelectedIndex && t.Text == _view.ActiveMarkerText).Any();

            _view.SaveEnabled = !markerExists;
            _view.AddEnabled = !markerExists;
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
    }
}
