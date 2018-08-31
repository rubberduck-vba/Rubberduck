using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.CodeAnalysis.CodeMetrics;

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
    }
}
