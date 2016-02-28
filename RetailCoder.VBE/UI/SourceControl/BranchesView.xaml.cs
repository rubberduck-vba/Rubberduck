using Ninject;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for BranchesView.xaml
    /// </summary>
    public partial class BranchesView : IControlView
    {
        public BranchesView()
        {
            InitializeComponent();
        }

        [Inject]
        public BranchesView(IControlViewModel vm) : this()
        {
            DataContext = vm;
        }

        public IControlViewModel ViewModel { get { return (IControlViewModel)DataContext; } }
    }
}
