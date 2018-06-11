using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameView : IRefactoringView
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
