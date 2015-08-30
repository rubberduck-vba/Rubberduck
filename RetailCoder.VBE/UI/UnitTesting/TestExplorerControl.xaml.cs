using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interactivity;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;

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

        private void TestExplorerControl_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var oldValue = e.OldValue as TestExplorerViewModel;
            if (oldValue != null)
            {
                oldValue.TestCompleted -= ViewModel_TestCompleted;
            }

            var newValue = e.NewValue as TestExplorerViewModel;
            if (newValue == null)
            {
                return;
            }

            newValue.TestCompleted += ViewModel_TestCompleted;
        }

        private void ViewModel_TestCompleted(object sender, TestCompletedEventArgs e)
        {
            UpdateTestMethod(e.Test);
        }

        private TestExplorerViewModel Context { get { return DataContext as TestExplorerViewModel; } }

        private void UpdateTestMethod(TestMethod test)
        {
            var view = ((CollectionViewSource)Resources["OutcomeGroupViewSource"]).View;
            view.Refresh();
        }

        private void TreeView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (Context == null)
            {
                return;
            }

            var selection = Context.SelectedItem;
            if (selection == null)
            {
                return;
            }

            Context.NavigateCommand.Execute(selection.GetNavigationArgs());
        }
    }

    /// <summary>
    /// http://stackoverflow.com/a/5118406/1188513
    /// </summary>
    public class BindableSelectedItemBehavior : Behavior<TreeView>
    {
        public object SelectedItem
        {
            get { return (object) GetValue(SelectedItemProperty); }
            set { SetValue(SelectedItemProperty, value); }
        }

        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register("SelectedItem", typeof (object), typeof (BindableSelectedItemBehavior),
                new UIPropertyMetadata(null, OnSelectedItemChanged));

        private static void OnSelectedItemChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            var item = e.NewValue as TreeViewItem;
            if (item != null)
            {
                item.SetValue(TreeViewItem.IsSelectedProperty, true);
            }
        }

        protected override void OnAttached()
        {
            base.OnAttached();

            this.AssociatedObject.SelectedItemChanged += OnTreeViewSelectedItemChanged;
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();

            if (this.AssociatedObject != null)
            {
                this.AssociatedObject.SelectedItemChanged -= OnTreeViewSelectedItemChanged;
            }
        }

        private void OnTreeViewSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            this.SelectedItem = e.NewValue;
        }
    }
}
