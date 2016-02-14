using System.Windows;
using System.Windows.Input;

namespace Rubberduck.UI.Controls
{
    public partial class GroupingGrid
    {
        public static readonly DependencyProperty IsExpandedProperty =
            DependencyProperty.Register("IsExpanded", typeof (bool), typeof (GroupingGrid));

        public static readonly DependencyProperty ShowGroupingItemCountProperty =
            DependencyProperty.Register("ShowGroupingItemCount", typeof (bool), typeof (GroupingGrid));

        public bool IsExpanded
        {
            get { return (bool)GetValue(IsExpandedProperty); }
            set { SetValue(IsExpandedProperty, value); }
        }

        public bool ShowGroupingItemCount
        {
            get { return (bool) GetValue(ShowGroupingItemCountProperty); }
            set { SetValue(ShowGroupingItemCountProperty, value); }
        }

        public GroupingGrid()
        {
            InitializeComponent();
        }

        private void GroupingGridItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var context = DataContext as INavigateSelection;
            if (context != null)
            {
                var selection = context.SelectedItem;
                context.NavigateCommand.Execute(selection.GetNavigationArgs());
            }
        }
    }
}
