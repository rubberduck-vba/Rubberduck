using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveCloserToUsage;
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

namespace Rubberduck.UI.Refactorings.MoveCloserToUsage
{
    /// <summary>
    /// Interaktionslogik für MoveCloserToUsageView.xaml
    /// </summary>
    public partial class MoveCloserToUsageView : IRefactoringView<MoveCloserToUsageModel>
    {
        public MoveCloserToUsageView()
        {
            InitializeComponent();
        }

    }
}
