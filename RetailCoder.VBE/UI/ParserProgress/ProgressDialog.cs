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
            parserProgessControl.DataContext = viewModel;
            parserProgessControl.ExpanderStateChanged += parserProgessControl_ExpanderStateChanged;

            parser.ParseCompleted += parser_ParseCompleted;
        }

        void parserProgessControl_ExpanderStateChanged(object sender, ParserProgessControl.ExpanderStateChangedEventArgs e)
        {
            Height = e.IsExpanded ? 200 : 96;
        }

        void parser_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            Result = e.ParseResults.FirstOrDefault();
        }

        //public for designer only
        public ProgressDialog()
        {
            InitializeComponent();
        }

        public VBProjectParseResult Result { get; private set; }
    }
}
