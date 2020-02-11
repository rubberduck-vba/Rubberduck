using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;

namespace Rubberduck.UI.Refactorings.MoveMember
{
    /// <summary>
    /// Interaction logic for MoveMemberView.xaml
    /// </summary>
    public partial class MoveMemberView : IRefactoringView<MoveMemberModel> //UserControl
    {
        public MoveMemberView()
        {
            InitializeComponent();
        }
    }
}
