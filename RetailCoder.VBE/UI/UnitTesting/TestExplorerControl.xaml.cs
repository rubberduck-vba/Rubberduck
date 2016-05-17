using System;
using System.Windows.Data;
using System.Windows.Threading;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// Interaction logic for TestExplorerControl.xaml
    /// </summary>
    public partial class TestExplorerControl
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
            _dispatcher.Invoke(UpdateUI);
        }

        private void UpdateUI()
        {
            var resource = FindResource("ResultsByOutcome") as CollectionViewSource;
            if (resource != null)
            {
                resource.View.Refresh();
            }
        }
    }
}
