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
        public ProgressDialog(IRubberduckParser parser, VBProject project)
            : this()
        {
            var viewModel = new ParserProgessViewModel(parser, project);
            viewModel.Completed += viewModel_Completed;

            parserProgessControl.DataContext = viewModel;
            parserProgessControl.ExpanderStateChanged += parserProgessControl_ExpanderStateChanged;
        }

        void viewModel_Completed(object sender, ParseCompletedEventArgs e)
        {
            Result = e.ParseResults.FirstOrDefault();
            Close();
        }

        void parserProgessControl_ExpanderStateChanged(object sender, ParserProgessControl.ExpanderStateChangedEventArgs e)
        {
            Height = e.IsExpanded ? 255 : 96;
        }

        //public for designer only
        public ProgressDialog()
        {
            InitializeComponent();
        }

        public VBProjectParseResult Result { get; private set; }
    }
}
