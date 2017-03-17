using System;
using System.Linq;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for InspectionSettings.xaml
    /// </summary>
    public partial class InspectionSettings : ISettingsView
    {
        public InspectionSettings()
        {
            InitializeComponent();
        }

        public InspectionSettings(ISettingsViewModel vm) : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel { get { return DataContext as ISettingsViewModel; } }

        private void GroupingGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Cancel || e.EditAction == DataGridEditAction.Cancel) { return; }
            
            var selectedSeverityName = ((ComboBox) e.EditingElement).SelectedItem.ToString();

            var severities = Enum.GetValues(typeof(CodeInspectionSeverity)).Cast<CodeInspectionSeverity>();
            var selectedSeverity = severities.Single(s => RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + s, Settings.Culture) == selectedSeverityName);

            ((InspectionSettingsViewModel) ViewModel).UpdateCollection(selectedSeverity);
        }

        private void WhitelistedIdentifierGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Cancel || e.EditAction == DataGridEditAction.Cancel) { return; }

            var identifiers = WhitelistedIdentifiersGrid.ItemsSource.OfType<WhitelistedIdentifierSetting>().ToList();

            var editedIndex = e.Row.GetIndex();
            identifiers.RemoveAt(editedIndex);
            identifiers.Insert(editedIndex, new WhitelistedIdentifierSetting(((TextBox)e.EditingElement).Text));

            ((InspectionSettingsViewModel)ViewModel).WhitelistedIdentifierSettings = new ObservableCollection<WhitelistedIdentifierSetting>(identifiers);
        }

        private void AddNewItem(object sender, System.Windows.RoutedEventArgs e)
        {
            WhitelistedIdentifiersGrid.CommitEdit();
            ((InspectionSettingsViewModel) ViewModel).AddWhitelistedNameCommand.Execute(null);
            e.Handled = true;
        }
    }
}
