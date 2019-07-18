using Rubberduck.Interaction.Navigation;
using System.Windows;
using System.Windows.Input;

namespace Rubberduck.UI.Controls
{
    public partial class GroupingGrid
    {
        public static readonly DependencyProperty ShowGroupingItemCountProperty =
            DependencyProperty.Register(nameof(ShowGroupingItemCount), typeof (bool), typeof(GroupingGrid));

        public static readonly DependencyProperty InitialExpandedStateProperty =
            DependencyProperty.Register(nameof(InitialExpandedState), typeof(bool), typeof(GroupingGrid));

        public bool ShowGroupingItemCount
        {
            get => (bool)GetValue(ShowGroupingItemCountProperty);
            set => SetValue(ShowGroupingItemCountProperty, value);
        }

        public bool InitialExpandedState
        {
            get => (bool)GetValue(InitialExpandedStateProperty);
            set => SetValue(InitialExpandedStateProperty, value);
        }

        public GroupingGrid()
        {
            InitializeComponent();
        }

        private void GroupingGridItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var context = DataContext as INavigateSelection;
            var selection = context?.SelectedItem;

            if (selection != null)
            {
                context.NavigateCommand.Execute(selection.GetNavigationArgs());
            }
        }
    }
}
