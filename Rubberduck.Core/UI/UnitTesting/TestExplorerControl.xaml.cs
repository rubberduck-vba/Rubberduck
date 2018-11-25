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
                if (FindResource("ResultsByOutcome") is CollectionViewSource outcome)
                {
                    outcome.View.Refresh();
                }
                if (FindResource("ResultsByModule") is CollectionViewSource module)
                {
                    module.View.Refresh();
                }
                if (FindResource("ResultsByCategory") is CollectionViewSource category)
                {
                    category.View.Refresh();
                }
            });
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed || DataContext == null) { return; }

            ((TestExplorerViewModel)DataContext).TestCompleted -= OnTestCompleted;
            DataContextChanged -= TestExplorerControl_DataContextChanged;

            _isDisposed = true;
        }
    }
}
