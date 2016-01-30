using System.Windows;

namespace Rubberduck.UI.Controls.GroupingGrid
{
    public partial class GroupingGrid 
    {
        public static readonly DependencyProperty IsExpandedProperty =
       DependencyProperty.Register("IsExpanded", typeof(bool), typeof(GroupingGrid));

        public bool IsExpanded
        {
            get { return (bool)GetValue(IsExpandedProperty); }
            set { SetValue(IsExpandedProperty, value); }
        }

        public GroupingGrid()
        {
            InitializeComponent();
        }
    }
}
