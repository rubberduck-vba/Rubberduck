using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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

        public InspectionResultsControl()
        {
            InitializeComponent();

            _inspectionTypeGroupsViewSource = (CollectionViewSource)FindResource("InspectionTypeGroupViewSource");
            _inspectionTypeGroupsTemplate = (DataTemplate)FindResource("InspectionTypeGroupsTemplate");

            _moduleGroupsViewSource = (CollectionViewSource)FindResource("CodeModuleGroupViewSource");
            _moduleGroupsTemplate = (DataTemplate)FindResource("CodeModuleGroupsTemplate");
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
    }
}
