using System;
using System.Linq;
using System.Windows.Controls;
using Rubberduck.Inspections;

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
            var selectedSeverity = severities.Single(s => RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + s) == selectedSeverityName);

            ((InspectionSettingsViewModel) ViewModel).UpdateCollection(selectedSeverity);
        }
    }
}
