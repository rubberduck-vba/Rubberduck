namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameView
    {
        public RenameView()
        {
            InitializeComponent();

            base.Loaded += delegate
                {
                    RenameTextBox.Focus();
                    RenameTextBox.SelectAll();
                };
        }
    }
}
