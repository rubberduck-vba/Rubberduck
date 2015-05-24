using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI
{
    public partial class ProgressDialog : Form
    {
        private readonly IRubberduckParser _parser;
        private readonly BackgroundWorker _bgw = new BackgroundWorker();
        private readonly VBProject _project;

        public ProgressDialog(IRubberduckParser parser, VBProject project)
            : this()
        {
            _parser = parser;
            _project = project;

            Shown += ProgressDialog_Shown;
            _bgw.WorkerReportsProgress = true;
            _bgw.DoWork += _bgw_DoWork;
            _bgw.RunWorkerCompleted += _bgw_RunWorkerCompleted;
        }

        public ProgressDialog()
        {
            InitializeComponent();
        }

        public VBProjectParseResult Result { get; private set; }

        private void ProgressDialog_Shown(object sender, EventArgs e)
        {
            _bgw.RunWorkerAsync();
        }

        private void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParseProgress += _parser_ParseProgress;
            Result = _parser.Parse(_project);
            _parser.ParseStarted -= _parser_ParseStarted;
            _parser.ParseProgress -= _parser_ParseProgress;
        }

        private void _bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Close();
        }

        void _parser_ParseProgress(object sender, ResolutionProgressEventArgs e)
        {
            SetStatus("Resolving '" + e.ParseResult.QualifiedName.ComponentName + "'...");
        }

        void _parser_ParseStarted(object sender, ParseStartedEventArgs e)
        {
            SetStatus("Parsing project components...");
        }

        public void SetStatus(string status)
        {
            Invoke(((MethodInvoker) delegate
            {
                TitleLabel.Text = status;
                Refresh();
            }));
        }
    }
}
