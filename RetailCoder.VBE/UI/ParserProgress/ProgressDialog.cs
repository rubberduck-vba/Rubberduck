using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;

namespace Rubberduck.UI.ParserProgress
{
    public partial class ProgressDialog : Form
    {
        private readonly int _initialSize;
        private const int ExpandedSize = 255;

        public ProgressDialog(ParserProgessViewModel viewModel)
            : this()
        {
            _initialSize = Height;

            viewModel.Completed += viewModel_Completed;

            parserProgessControl.DataContext = viewModel;
            parserProgessControl.ExpanderStateChanged += parserProgessControl_ExpanderStateChanged;
        }

        //public for designer only
        public ProgressDialog()
        {
            InitializeComponent();
        }

        private void viewModel_Completed(object sender, ParseCompletedEventArgs e)
        {
            Result = e.ParseResults.FirstOrDefault();
            if (InvokeRequired)
            {
                BeginInvoke((MethodInvoker) Hide);
            }
            else
            {
                Hide();
            }
        }

        private void parserProgessControl_ExpanderStateChanged(object sender, ParserProgessControl.ExpanderStateChangedEventArgs e)
        {
            Height = e.IsExpanded ? ExpandedSize : _initialSize;
        }

        public VBProjectParseResult Result { get; private set; }
    }
}
