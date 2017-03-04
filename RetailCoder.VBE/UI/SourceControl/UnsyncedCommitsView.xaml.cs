namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for UnsyncedCommitsView.xaml
    /// </summary>
    public partial class UnsyncedCommitsView : IControlView
    {
        public UnsyncedCommitsView()
        {
            InitializeComponent();
        }

        public UnsyncedCommitsView(IControlViewModel vm) : this()
        {
            DataContext = vm;
        }

        public IControlViewModel ViewModel { get { return (IControlViewModel)DataContext; } }
    }
}
