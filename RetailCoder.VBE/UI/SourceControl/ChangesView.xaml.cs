namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for ChangesView.xaml
    /// </summary>
    public partial class ChangesView : IControlView
    {
        public ChangesView()
        {
            InitializeComponent();
        }

        public ChangesView(IControlViewModel vm) : this()
        {
            DataContext = vm;
        }

        public IControlViewModel ViewModel { get { return (IControlViewModel)DataContext; } }
    }
}
