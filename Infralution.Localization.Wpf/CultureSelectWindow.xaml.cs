using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Globalization;
namespace Infralution.Localization.Wpf
{
    /// <summary>
    /// Window that allows the user to select the culture to use at design time
    /// </summary>
    public partial class CultureSelectWindow : Window
    {

        /// <summary>
        /// Handle sorting Culture Info
        /// </summary>
        private class CultureInfoComparer : Comparer<CultureInfo>        
        {
            public override int Compare(CultureInfo x, CultureInfo y)
            {
                return x.DisplayName.CompareTo(y.DisplayName);
            }
        }

        /// <summary>
        /// Create a new instance of the window
        /// </summary>
        public CultureSelectWindow()
        {
            InitializeComponent();
            List<CultureInfo> cultures = new List<CultureInfo>(CultureInfo.GetCultures(CultureTypes.SpecificCultures));
            cultures.Sort(new CultureInfoComparer());
            _cultureCombo.ItemsSource = cultures;
            _cultureCombo.SelectedItem = CultureManager.UICulture;
        }


        /// <summary>
        /// Set the CultureManager.UICulture
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _cultureCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CultureManager.UICulture = _cultureCombo.SelectedItem as CultureInfo;
        }
    }
}
