using System;
using System.Windows.Data;
using System.Windows.Threading;

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

        void TestExplorerControl_DataContextChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
        {
            var oldContext = e.OldValue as TestExplorerViewModel;
            if (oldContext != null)
            {
                oldContext.TestCompleted -= OnTestCompleted;
            }

            var context = e.NewValue as TestExplorerViewModel;
            if (context != null)
            {
                context.TestCompleted += OnTestCompleted;
            }
        }

        private void OnTestCompleted(object sender, EventArgs eventArgs)
        {
            _dispatcher.Invoke(() =>
            {
                var resource = FindResource("ResultsByOutcome") as CollectionViewSource;
                if (resource != null)
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
    }
}
