using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;

namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameView : IRefactoringView<RenameModel>
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
