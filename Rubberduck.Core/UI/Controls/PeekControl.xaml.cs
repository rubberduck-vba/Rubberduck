using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// Interaction logic for PeekControl.xaml
    /// </summary>
    public partial class PeekControl : UserControl
    {
        public PeekControl()
        {
            InitializeComponent();
            DragThumb.DragDelta += DragThumb_DragDelta;
        }

        private void DragThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (this.Parent is Popup parent)
            {
                parent.HorizontalOffset += e.HorizontalChange;
                parent.VerticalOffset += e.VerticalChange;
                e.Handled = true;
            }
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            try
            {
                if (e.ChangedButton == MouseButton.Left)
                {
                    DragThumb.RaiseEvent(e);
                    e.Handled = true;
                }
            }
            catch 
            {
                // possible stack overflow exception with rapid double-button clickety-clicking
            }
        }

        private void LinkButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is PeekDefinitionViewModel vm)
            {
                vm.CloseCommand.Execute(null);
            }
        }
    }
}
