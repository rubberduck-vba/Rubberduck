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
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private IEnumerable<VBProjectParseResult> _parseResults;
        private IList<ICodeInspectionResult> _results;
        private GridViewSort<CodeInspectionResultGridViewItem> _gridViewSort;
        private readonly IInspector _inspector;

        /// <summary>
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown when <see cref="_inspector_Reset"/> is <c>null</c>.</exception>
        /// <param name="inspector"></param>
        /// <param name="vbe"></param>
        /// <param name="addin"></param>
        /// <param name="window"></param>
        public CodeInspectionsDockablePresenter(IInspector inspector, VBE vbe, AddIn addin, CodeInspectionsWindow window, GridViewSort<CodeInspectionResultGridViewItem> gridViewSort)
            :base(vbe, addin, window)
        {
            _inspector = inspector;
            _inspector.IssuesFound += _inspector_IssuesFound;
            _inspector.Reset += _inspector_Reset;
            _inspector.Parsing += _inspector_Parsing;
            _inspector.ParseCompleted += _inspector_ParseCompleted;

            _gridViewSort = gridViewSort;

            Control.RefreshCodeInspections += Control_RefreshCodeInspections;
            Control.NavigateCodeIssue += Control_NavigateCodeIssue;
            Control.QuickFix += Control_QuickFix;
            Control.CopyResults += Control_CopyResultsToClipboard;
            Control.Cancel += Control_Cancel;
            Control.SortColumn += SortColumn;
        }

        private void SortColumn(object sender, DataGridViewCellMouseEventArgs e)
        {
            var columnName = Control.GridView.Columns[e.ColumnIndex].Name;
            if (columnName == "Icon") { columnName = "Severity"; }

            Control.InspectionResults = new BindingList<CodeInspectionResultGridViewItem>(_gridViewSort.Sort(Control.InspectionResults.AsEnumerable(), columnName).ToList());
        }

        private void Control_Cancel(object sender, EventArgs e)
        {
            if (_cancelTokenSource != null)
            { 
                _cancelTokenSource.Cancel();
            }
        }

        private void _inspector_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            if (sender != this)
            {
                return;
            }

            ToggleParsingStatus(false);
            _parseResults = e.ParseResults;
        }

        private void _inspector_Parsing(object sender, EventArgs e)
        {
            if (sender != this)
            {
                return;
            }

            ToggleParsingStatus();
            Control.Invoke((MethodInvoker) delegate
            {
                Control.EnableRefresh(false);
            });
        }

        private void ToggleParsingStatus(bool isParsing = true)
        {
            Control.Invoke((MethodInvoker) delegate
            {
                Control.ToggleParsingStatus(isParsing);
            });
        }

        private void Control_CopyResultsToClipboard(object sender, EventArgs e)
        {
            var results = string.Join("\n", _results.Select(FormatResultForClipboard));
            var resource = _results.Count == 1
                ? RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;
            var text = string.Format(resource, DateTime.Now, _results.Count) + results;

            Clipboard.SetText(text);
        }

        private string FormatResultForClipboard(ICodeInspectionResult result)
        {
            var module = result.QualifiedSelection.QualifiedName;
            return string.Format(
                "{0}: {1} - {2}.{3}, line {4}",
                result.Severity,
                result.Name,
                module.ProjectName,
                module.ComponentName,
                result.QualifiedSelection.Selection.StartLine);
        }

        private int _issues;
        private void _inspector_IssuesFound(object sender, InspectorIssuesFoundEventArg e)
        {
            Interlocked.Add(ref _issues, e.Issues.Count);
            Control.Invoke((MethodInvoker) delegate
            {
                var newCount = _issues;
                Control.SetIssuesStatus(newCount);
            });
        }

        private void Control_QuickFix(object sender, QuickFixEventArgs e)
        {
            e.QuickFix();
            Control_RefreshCodeInspections(null, EventArgs.Empty);
        }

        public override void Show()
        {
            base.Show();
            Refresh();
        }

        private void Control_NavigateCodeIssue(object sender, NavigateCodeEventArgs e)
        {
            try
            {
                if (e.QualifiedName.Component == null)
                {
                    return;
                }
                e.QualifiedName.Component.CodeModule.CodePane.SetSelection(e.Selection);
            }
            catch (COMException)
            {
                // gulp
            }
        }

        private void Control_RefreshCodeInspections(object sender, EventArgs e)
        {
            Refresh();
        }

        private CancellationTokenSource _cancelTokenSource;
        private async void Refresh()
        {
            _cancelTokenSource = new CancellationTokenSource();
            var token = _cancelTokenSource.Token;

            Control.EnableRefresh(false);
            Control.Cursor = Cursors.WaitCursor;

            try
            {
                await Task.Run(() => RefreshAsync(token), token);
                if (_results != null)
                {
                    var results = _results.Select(item => new CodeInspectionResultGridViewItem(item));

                    Control.SetContent(new BindingList<CodeInspectionResultGridViewItem>(
                    _gridViewSort.Sort(results, _gridViewSort.ColumnName,
                        _gridViewSort.SortedAscending).ToList()));
                }
            }
            catch (TaskCanceledException)
            {
            }
            finally
            {
                Control.SetIssuesStatus(_issues, true);
                Control.EnableRefresh();
                Control.Cursor = Cursors.Default;
            }
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
            Control.Invoke((MethodInvoker) delegate
            {
                Control.SetIssuesStatus(_issues);
                Control.InspectionResults.Clear();
                Control.EnableRefresh();
                Control.Cursor = Cursors.Default;
            });
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            _inspector.IssuesFound -= _inspector_IssuesFound;
            _inspector.Reset -= _inspector_Reset;
            _inspector.Parsing -= _inspector_Parsing;
            _inspector.ParseCompleted -= _inspector_ParseCompleted;

            Control.RefreshCodeInspections -= Control_RefreshCodeInspections;
            Control.NavigateCodeIssue -= Control_NavigateCodeIssue;
            Control.QuickFix -= Control_QuickFix;
            Control.CopyResults -= Control_CopyResultsToClipboard;
            Control.Cancel -= Control_Cancel;
            Control.SortColumn -= SortColumn;

            base.Dispose(true);
        }
    }
}
