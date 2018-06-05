using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Rubberduck.UI.Settings
{
    /// <summary>
    /// Interaction logic for AutoCompleteSettings.xaml
    /// </summary>
    public partial class AutoCompleteSettings : UserControl, ISettingsView
    {
        public AutoCompleteSettings()
        {
            InitializeComponent();
        }

        public AutoCompleteSettings(ISettingsViewModel vm) : this()
        {
            DataContext = vm;
        }

        public ISettingsViewModel ViewModel => DataContext as ISettingsViewModel;
    }
}
