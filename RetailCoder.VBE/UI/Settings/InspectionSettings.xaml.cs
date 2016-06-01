using System.Windows.Controls;

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

            ((InspectionSettingsViewModel) ViewModel).UpdateCollection();
        }
    }
}
