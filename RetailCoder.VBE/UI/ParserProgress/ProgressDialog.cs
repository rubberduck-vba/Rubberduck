using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;

namespace Rubberduck.UI.ParserProgress
{
    public partial class ProgressDialog : Form
    {
        private readonly ParserProgessViewModel _viewModel;
        private readonly int _initialSize;
        private const int ExpandedSize = 255;

        public ProgressDialog(ParserProgessViewModel viewModel)
            : this()
        {
            _viewModel = viewModel;
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

        private void viewModel_Completed(object sender, EventArgs e)
        {
            Result = _viewModel.Parser.Declarations;
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
