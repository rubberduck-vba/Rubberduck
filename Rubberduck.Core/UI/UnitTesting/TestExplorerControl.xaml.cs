using System;
using System.Windows.Data;
using System.Windows.Threading;
using Rubberduck.UI.Controls;

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

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing || DataContext == null)
            {
                return;
            }

            ((TestExplorerViewModel)DataContext).TestCompleted -= OnTestCompleted;
            DataContextChanged -= TestExplorerControl_DataContextChanged;
            _isDisposed = true;
        }
        private GroupingGrid _newGroupingGrid;
        private void ResultsGroupingGrid_IsVisibleChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
        {
            if (sender is GroupingGrid)
            {
                GroupingGrid groupingGrid = (GroupingGrid)sender;
                // actual height checks that form is visible at this point
                // don't set up the "new" visible element until we see an object with an actual height
                if (!((groupingGrid.ActualHeight == 0) && (_newGroupingGrid is null))) // check form elements are actually set up (i.e. not initialising form)
                {
                    // this code relies on the new element being made visible before the old is hidden, so we save the new element 
                    // and then copy column widths from the old now hidden element on the subsequent invocation
                    if (groupingGrid.IsVisible)
                        _newGroupingGrid = groupingGrid;
                    else
                    {
                        if (!(_newGroupingGrid is null))
                        {
                            // copy column widths from previously visible grouping grid to ensure widths are consistent
                            for (int columnIndex = 0; columnIndex < groupingGrid.Columns.Count; columnIndex++)
                            {
                                // copy column widths from current now hidden columns to the new visible columns
                                _newGroupingGrid.Columns[columnIndex].Width = groupingGrid.Columns[columnIndex].Width;
                            }
                        }
                    }
                }
            }
        }
    }
}
