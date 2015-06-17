using System;
using System.ComponentModel;
using System.Linq;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingPresenter
    {
        private readonly ITodoSettingsView _view;
        private readonly IAddTodoMarkerView _addTodoMarkerView;

        public ToDoMarker ActiveMarker
        {
            get { return _view.TodoMarkers[_view.SelectedIndex]; }
        }

        public TodoSettingPresenter(ITodoSettingsView view, IAddTodoMarkerView addTodoMarkerView)
        {
            _view = view;
            _addTodoMarkerView = addTodoMarkerView;

            _view.AddMarker += AddMarker;
            _view.RemoveMarker += RemoveMarker;
            _view.PriorityChanged += SaveMarker;

            _addTodoMarkerView.AddMarker += ConfirmAddMarker;
            _addTodoMarkerView.Cancel += CancelAddMarker;
            _addTodoMarkerView.TextChanged += AddMarkerTextChanged;
        }

        ~TodoSettingPresenter()
        {
            _addTodoMarkerView.Close();
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
            _addTodoMarkerView.TodoMarkers = _view.TodoMarkers.ToList();
            _addTodoMarkerView.Show();
        }

        private void AddMarkerTextChanged(object sender, EventArgs e)
        {
            _addTodoMarkerView.IsValidMarker = _view.TodoMarkers.All(t => t.Text != _addTodoMarkerView.MarkerText.ToUpper()) && 
                                               _addTodoMarkerView.MarkerText != string.Empty;
        }

        private void ConfirmAddMarker(object sender, EventArgs e)
        {
            HideAddMarkerForm();
            _view.TodoMarkers = new BindingList<ToDoMarker>(_addTodoMarkerView.TodoMarkers);
        }

        private void CancelAddMarker(object sender, EventArgs e)
        {
            HideAddMarkerForm();
        }

        private void HideAddMarkerForm()
        {
            _addTodoMarkerView.Hide();
            _addTodoMarkerView.MarkerText = string.Empty;
            _addTodoMarkerView.MarkerPriority = default(TodoPriority);
        }

        public void SetActiveItem(int index)
        {
            _view.SelectedIndex = index;
        }
    }
}
