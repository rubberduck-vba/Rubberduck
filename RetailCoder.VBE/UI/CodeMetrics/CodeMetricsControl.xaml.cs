using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.Navigation.CodeMetrics;

namespace Rubberduck.UI.CodeMetrics
{
    /// <summary>
    /// Interaction logic for CodeMetricsControl.xaml
    /// </summary>
    public partial class CodeMetricsControl
    {
        public CodeMetricsControl()
        {
            InitializeComponent();
        }

        private CodeMetricsViewModel ViewModel { get { return DataContext as CodeMetricsViewModel; } }
        private void SearchBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            ViewModel.FilterByName(ViewModel.ModuleMetrics, ((TextBox)sender).Text);
        }

        private void SearchIcon_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            SearchBox.Focus();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            ClearSearchBox();
        }

        private void ClearSearchBox()
        {
            SearchBox.Text = string.Empty;
            SearchBox.Focus();
        }

        private void SearchBox_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                ClearSearchBox();
            }
        }
    }
}
