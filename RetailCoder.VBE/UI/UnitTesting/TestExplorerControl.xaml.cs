using System;
using System.Windows.Controls;
using System.Windows.Data;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// Interaction logic for TestExplorerControl.xaml
    /// </summary>
    public partial class TestExplorerControl : UserControl
    {
        public TestExplorerControl()
        {
            InitializeComponent();
            DataContextChanged += TestExplorerControl_DataContextChanged;
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
            try
            {
                var resource = FindResource("ResultsByOutcome") as CollectionViewSource;
                if (resource != null)
                {
                    resource.View.Refresh();
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
