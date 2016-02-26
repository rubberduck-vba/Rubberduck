using Ninject;

namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for UnsyncedCommitsView.xaml
    /// </summary>
    public partial class UnsyncedCommitsView
    {
        public UnsyncedCommitsView()
        {
            InitializeComponent();
        }
        [Inject]
        public UnsyncedCommitsView(UnsyncedCommitsViewViewModel vm) : this()
        {
            DataContext = vm;
        }
    }
}
