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

        public static readonly DependencyProperty GroupedItemSourceProperty =
            DependencyProperty.Register("GroupedItemSource", typeof (CollectionViewSource), typeof (GroupingGrid));

        public CollectionViewSource GroupedItemSource
        {
            get { return (CollectionViewSource) GetValue(GroupedItemSourceProperty); }
            set { SetValue(GroupedItemSourceProperty, value); }
        }

        public GroupingGrid()
        {
            InitializeComponent();
        }

        private void GroupingGridItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var context = DataContext as INavigateSelection;
            if (context == null)
            {
                return;
            }
            var selection = context.SelectedItem;
            if (selection != null)
            {
                context.NavigateCommand.Execute(selection.GetNavigationArgs());
            }
        }
    }
}
