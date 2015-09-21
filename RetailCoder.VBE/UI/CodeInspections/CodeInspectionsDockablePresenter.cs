using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.Parsing;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private IList<ICodeInspectionResult> _results;
        private readonly IInspector _inspector;

        /// <summary>
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown when <see cref="_inspector_Reset"/> is <c>null</c>.</exception>
        public CodeInspectionsDockablePresenter(IInspector inspector, VBE vbe, AddIn addin, CodeInspectionsWindow window)
            :base(vbe, addin, window)
        {
            _inspector = inspector;
            _inspector.IssuesFound += _inspector_IssuesFound;
            _inspector.Reset += _inspector_Reset;
            _inspector.Parsing += _inspector_Parsing;
            _inspector.ParseCompleted += _inspector_ParseCompleted;
        }

        private void _inspector_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            if (sender != this)
            {
                return;
            }

            //ToggleParsingStatus(false);
        }

        private void _inspector_Parsing(object sender, EventArgs e)
        {
            if (sender != this)
            {
                return;
            }

            //ToggleParsingStatus();
        }

        private void Control_CopyResultsToClipboard(object sender, EventArgs e)
        {
        }

        private int _issues;
        private void _inspector_IssuesFound(object sender, InspectorIssuesFoundEventArg e)
        {
            Interlocked.Add(ref _issues, e.Issues.Count);
        }

        public override void Show()
        {
            base.Show();
            Refresh();
        }

        private CancellationTokenSource _cancelTokenSource;
        private async void Refresh()
        {
            _cancelTokenSource = new CancellationTokenSource();
            var token = _cancelTokenSource.Token;

            Control.Cursor = Cursors.WaitCursor;

            await Task.Run(() => RefreshAsync(token), token);
            if (_results != null)
            {
                var results = _results.Select(item => new CodeInspectionResultGridViewItem(item));

                //Control.SetContent(new BindingList<CodeInspectionResultGridViewItem>(
                //_gridViewSort.Sort(results, _gridViewSort.ColumnName,
                //    _gridViewSort.SortedAscending).ToList();
            }

            Control.Cursor = Cursors.Default;
        }

        private async Task RefreshAsync(CancellationToken token)
        {
            try
            {
                var projectParseResult = await _inspector.Parse(VBE.ActiveVBProject, this);
                _results = await _inspector.FindIssuesAsync(projectParseResult, token);
            }
            catch (TaskCanceledException)
            {
                // If FindIssuesAsync is canceled, we can leave the old results or 
                // create a new List. Let's leave the old ones for now.
            }
            catch (COMException)
            {
                // burp
            }
        }

        private void _inspector_Reset(object sender, EventArgs e)
        {
            _issues = 0;
            //Control.Invoke((MethodInvoker) delegate
            //{
            //    Control.SetIssuesStatus(_issues);
            //    Control.InspectionResults.Clear();
            //    Control.EnableRefresh();
            //    Control.Cursor = Cursors.Default;
            //});
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            _inspector.IssuesFound -= _inspector_IssuesFound;
            _inspector.Reset -= _inspector_Reset;
            _inspector.Parsing -= _inspector_Parsing;
            _inspector.ParseCompleted -= _inspector_ParseCompleted;

            base.Dispose(true);
        }
    }
}
