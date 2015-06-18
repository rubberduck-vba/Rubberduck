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

        public TodoSettingPresenter(ITodoSettingsView view, IAddTodoMarkerView addTodoMarkerView)
        {
            _view = view;
            _addTodoMarkerView = addTodoMarkerView;

            _view.AddMarker += AddMarker;
            _view.RemoveMarker += RemoveMarker;
            _view.PriorityChanged += PriorityChanged;

            _addTodoMarkerView.AddMarker += ConfirmAddMarker;
            _addTodoMarkerView.Cancel += CancelAddMarker;
            _addTodoMarkerView.TextChanged += AddMarkerTextChanged;
        }

        ~TodoSettingPresenter()
        {
            _addTodoMarkerView.Close();
        }

        private void PriorityChanged(object sender, EventArgs e)
        {
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
            _addTodoMarkerView.IsValidMarker = _view.TodoMarkers.All(t => t.Text.Equals(_addTodoMarkerView.MarkerText, StringComparison.CurrentCultureIgnoreCase) && 
                                               _addTodoMarkerView.MarkerText != string.Empty);
        }

        private void ConfirmAddMarker(object sender, EventArgs e)
        {
            _addTodoMarkerView.TodoMarkers.Add(new ToDoMarker(_addTodoMarkerView.MarkerText,
                _addTodoMarkerView.MarkerPriority));
            _view.TodoMarkers = new BindingList<ToDoMarker>(_addTodoMarkerView.TodoMarkers);
            HideAddMarkerForm();
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
    }
}
