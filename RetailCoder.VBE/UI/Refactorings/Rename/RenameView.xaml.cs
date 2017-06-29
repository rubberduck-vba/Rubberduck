namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameView
    {
        public RenameView()
        {
            InitializeComponent();

            Loaded +=
                {
                    RenameTextBox.Focus();
                    RenameTextBox.SelectAll();
                };
        }
    }
}
