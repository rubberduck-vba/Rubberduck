using System;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using System.Windows.Input;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// Interaction logic for TestExplorerControl.xaml
    /// </summary>
    public partial class TestExplorerControl : IDisposable
    {
        private readonly Dispatcher _dispatcher;

        public TestExplorerControl()
        {
            InitializeComponent();
            DataContextChanged += TestExplorerControl_DataContextChanged;
            _dispatcher = Dispatcher.CurrentDispatcher;
        }

        private void TestExplorerControl_DataContextChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
        {
            if (e.OldValue is TestExplorerViewModel oldContext)
            {
                oldContext.TestCompleted -= OnTestCompleted;
            }

            if (e.NewValue is TestExplorerViewModel context)
            {
                context.TestCompleted += OnTestCompleted;
            }
        }

        private void OnTestCompleted(object sender, EventArgs eventArgs)
        {
            _dispatcher.Invoke(() =>
            {
                if (FindResource("ResultsByOutcome") is CollectionViewSource resource)
                {
                    resource.View.Refresh();
                }
            });
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed || DataContext == null) { return; }

            DataContextChanged -= TestExplorerControl_DataContextChanged;
            ((TestExplorerViewModel)DataContext).TestCompleted -= OnTestCompleted;

            _isDisposed = true;
        }

        private void ScrollViewer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (e.Delta > 0)
            {
                ((ScrollViewer)sender).LineUp();
            }
            else
            {
                ((ScrollViewer)sender).LineDown();
            }

        }
    }
}
