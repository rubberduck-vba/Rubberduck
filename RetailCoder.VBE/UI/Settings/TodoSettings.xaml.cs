using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Controls;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for TodoSettings.xaml
    /// </summary>
    public partial class TodoSettings : ISettingsView
    {
        public TodoSettings()
        {
            InitializeComponent();
        }

        public TodoSettings(ISettingsViewModel vm)
            : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel { get { return DataContext as ISettingsViewModel; } }

        private void TodoMarkerGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Cancel || e.EditAction == DataGridEditAction.Cancel) { return; }

            var markers = TodoMarkerGrid.ItemsSource.OfType<ToDoMarker>().ToList();

            var editedIndex = e.Row.GetIndex();
            markers.RemoveAt(editedIndex);
            markers.Insert(editedIndex, new ToDoMarker(((TextBox)e.EditingElement).Text));

            ((TodoSettingsViewModel)ViewModel).TodoSettings = new ObservableCollection<ToDoMarker>(markers);
        }

        private void AddNewTodoMarker(object sender, System.Windows.RoutedEventArgs e)
        {
            TodoMarkerGrid.CommitEdit();
            ((TodoSettingsViewModel) ViewModel).AddTodoCommand.Execute(null);
            e.Handled = true;
        }
    }
}
