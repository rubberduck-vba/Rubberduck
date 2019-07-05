using System.Windows.Forms;
using UserControl = System.Windows.Controls.UserControl;

namespace Rubberduck.UI.Refactorings
{
    public partial class RefactoringDialogBase : Form
    {
        protected int MinWidth;
        protected int MinHeight;

        protected RefactoringDialogBase()
        {
            InitializeComponent();
            elementHost.AutoSize = true;
        }

        private UserControl _userControl;
        protected UserControl UserControl
        {
            get => _userControl;
            set
            {
                _userControl = value;
                elementHost.Child = _userControl;
                var proposedSize = elementHost.GetPreferredSize(new System.Drawing.Size(int.MaxValue, int.MaxValue));
                if (proposedSize.Height < MinHeight)
                {
                    proposedSize.Height = MinHeight;
                }
                if (proposedSize.Width < MinWidth)
                {
                    proposedSize.Width = MinWidth;
                }
                ClientSize = proposedSize;
            }
        }
    }
}
