using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings
{
    /// <summary>
    /// Interaction logic for ExtractInterfaceView.xaml
    /// </summary>
    public partial class ExtractInterfaceView : IRefactoringView<ExtractInterfaceModel>
    {
        public ExtractInterfaceView()
        {
            InitializeComponent();
        }
    }
}
