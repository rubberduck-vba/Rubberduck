using System;
using System.Windows.Input;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Interaction logic for ToDoExplorerControl.xaml
    /// </summary>
    public partial class ToDoExplorerControl
    {
        public ToDoExplorerControl()
        {
            InitializeComponent();
        }

        public event EventHandler TodoDoubleClick;
        private void GroupingGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var handler = TodoDoubleClick;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
