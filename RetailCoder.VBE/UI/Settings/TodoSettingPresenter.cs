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

            _view.AddMarker += AddMarker;
            _view.RemoveMarker += RemoveMarker;
            _view.PriorityChanged += SaveMarker;
        }

        private void SaveMarker(object sender, EventArgs e)
        {
            //todo: add test; How? I can't click the save button. Code smell here.
            var index = _view.SelectedIndex;
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

        public void SetActiveItem(int index)
        {
            _view.SelectedIndex = index;
        }
    }
}
