namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameView
    {
        public RenameView()
        {
            InitializeComponent();

            Loaded += (o, e) =>
                {
                    RenameTextBox.Focus();
                    RenameTextBox.SelectAll();
                };
        }
    }
}
