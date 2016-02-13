using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Rubberduck.Inspections;

namespace Rubberduck.UI.CodeInspections
{
    /// <summary>
    /// Interaction logic for InspectionResultsControl.xaml
    /// </summary>
    public partial class InspectionResultsControl : UserControl
    {
        private readonly CollectionViewSource _inspectionTypeGroupsViewSource;
        private readonly DataTemplate _inspectionTypeGroupsTemplate;

        private readonly CollectionViewSource _moduleGroupsViewSource;
        private readonly DataTemplate _moduleGroupsTemplate;

        private InspectionResultsViewModel ViewModel { get { return DataContext as InspectionResultsViewModel; } }

        public InspectionResultsControl()
        {
            InitializeComponent();

            _inspectionTypeGroupsViewSource = (CollectionViewSource)FindResource("InspectionTypeGroupViewSource");
            _inspectionTypeGroupsTemplate = (DataTemplate)FindResource("InspectionTypeGroupsTemplate");

            _moduleGroupsViewSource = (CollectionViewSource)FindResource("CodeModuleGroupViewSource");
            _moduleGroupsTemplate = (DataTemplate)FindResource("CodeModuleGroupsTemplate");

            Loaded += InspectionResultsControl_Loaded;
        }

        private void InspectionResultsControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (ViewModel.CanRefresh)
            {
                ViewModel.RefreshCommand.Execute(null);
            }
        }

        private bool _isModuleTemplate;
        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            _isModuleTemplate = TreeViewStyleToggle.IsChecked.HasValue && TreeViewStyleToggle.IsChecked.Value;
            InspectionResultsTreeView.ItemTemplate = _isModuleTemplate
                ? _moduleGroupsTemplate
                : _inspectionTypeGroupsTemplate;

            InspectionResultsTreeView.ItemsSource = _isModuleTemplate
                ? _moduleGroupsViewSource.View.Groups
                : _inspectionTypeGroupsViewSource.View.Groups;
        }

        private void InspectionResultsTreeView_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel == null || ViewModel.SelectedItem == null)
            {
                return;
            }

            var selectedResult = ViewModel.SelectedItem as CodeInspectionResultBase;
            if (selectedResult == null)
            {
                return;
            }

            var arg = selectedResult.QualifiedSelection.GetNavitationArgs();
            ViewModel.NavigateCommand.Execute(arg);
        }
    }
}
