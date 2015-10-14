using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI.ParserProgress
{
    public partial class ProgressDialog : Form
    {
        private readonly int _initialSize;
        private const int ExpandedSize = 255;

        public ProgressDialog(IRubberduckParser parser, VBProject project)
            : this()
        {
            _initialSize = Height;

            var viewModel = new ParserProgessViewModel(parser, project);
            viewModel.Completed += viewModel_Completed;

            parserProgessControl.DataContext = viewModel;
            parserProgessControl.ExpanderStateChanged += parserProgessControl_ExpanderStateChanged;
        }

        void viewModel_Completed(object sender, ParseCompletedEventArgs e)
        {
            Result = e.ParseResults.FirstOrDefault();
            Invoke((MethodInvoker) Hide);
        }

        void parserProgessControl_ExpanderStateChanged(object sender, ParserProgessControl.ExpanderStateChangedEventArgs e)
        {
            Height = e.IsExpanded ? ExpandedSize : _initialSize;
        }

        //public for designer only
        public ProgressDialog()
        {
            InitializeComponent();
        }

        public VBProjectParseResult Result { get; private set; }
    }
}
