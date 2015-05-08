using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.Parsing;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private IEnumerable<VBProjectParseResult> _parseResults;
        private IList<ICodeInspectionResult> _results;
        private readonly IInspector _inspector;

        public CodeInspectionsDockablePresenter(IInspector inspector, VBE vbe, AddIn addin, CodeInspectionsWindow window)
            :base(vbe, addin, window)
        {
            _inspector = inspector;
            _inspector.IssuesFound += OnIssuesFound;
            _inspector.Reset += OnReset;
            _inspector.Parsing += OnParsing;
            _inspector.ParseCompleted += OnParseCompleted;

            Control.RefreshCodeInspections += OnRefreshCodeInspections;
            Control.NavigateCodeIssue += OnNavigateCodeIssue;
            Control.QuickFix += OnQuickFix;
            Control.CopyResults += OnCopyResultsToClipboard;
        }

        private void OnParseCompleted(object sender, Parsing.ParseCompletedEventArgs e)
        {
            _parseResults = e.ParseResults;
            Task.Run(() => RefreshAsync());
        }

        private void OnParsing(object sender, EventArgs e)
        {
            Control.Invoke((MethodInvoker) delegate
            {
                Control.EnableRefresh(false);
            });
        }

        private void OnCopyResultsToClipboard(object sender, EventArgs e)
        {
            var results = string.Join("\n", _results.Select(FormatResultForClipboard));
            var text = string.Format("Rubberduck Code Inspections - {0}\n{1} issue" + (_results.Count != 1 ? "s" : string.Empty) + " found.\n",
                            DateTime.Now, _results.Count) + results;

            Clipboard.SetText(text);
        }

        private string FormatResultForClipboard(ICodeInspectionResult result)
        {
            var module = result.QualifiedSelection.QualifiedName;
            return string.Format(
                "{0}: {1} - {2}.{3}, line {4}",
                result.Severity,
                result.Name,
                module.Project.Name,
                module.Component.Name,
                result.QualifiedSelection.Selection.StartLine);
        }

        private void OnIssuesFound(object sender, InspectorIssuesFoundEventArg e)
        {
            Control.Invoke((MethodInvoker) delegate
            {
                var newCount = Control.IssueCount + e.Issues.Count;
                Control.IssueCount = newCount;
                Control.IssueCountText = string.Format("{0} issue" + (newCount != 1 ? "s" : string.Empty), newCount);
            });
        }

        private void OnQuickFix(object sender, QuickFixEventArgs e)
        {
            e.QuickFix(VBE);
            OnRefreshCodeInspections(null, EventArgs.Empty);
        }

        public override void Show()
        {
            base.Show();
            Task.Run(() => RefreshAsync());
        }

        private void OnNavigateCodeIssue(object sender, NavigateCodeEventArgs e)
        {
            try
            {
                e.QualifiedName.Component.CodeModule.CodePane.SetSelection(e.Selection);
            }
            catch (COMException)
            {
                // gulp
            }
        }

        private void OnRefreshCodeInspections(object sender, EventArgs e)
        {
            Task.Run(() => RefreshAsync());
        }

        private async Task RefreshAsync()
        {
            Control.Invoke((MethodInvoker) delegate
            {
                Control.EnableRefresh(false);
                Control.Cursor = Cursors.WaitCursor;
            });

            try
            {
                if (VBE != null)
                {
                    if (_parseResults == null)
                    {
                        _inspector.Parse(VBE);
                        return;
                    }

                    var parseResults = _parseResults.SingleOrDefault(p => p.Project == VBE.ActiveVBProject);
                    if (parseResults == null)
                    {
                        _inspector.Parse(VBE);
                        return;
                    }

                    _results = await _inspector.FindIssuesAsync(parseResults);

                    Control.Invoke((MethodInvoker) delegate
                    {
                        Control.SetContent(_results.Select(item => new CodeInspectionResultGridViewItem(item))
                            .OrderBy(item => item.Component)
                            .ThenBy(item => item.Line));
                    });
                }
            }
            catch (Exception exception)
            {
                // swallow
            }
            finally
            {
                Control.Invoke((MethodInvoker) delegate
                {
                    Control.Cursor = Cursors.Default;
                    Control.EnableRefresh();
                });
            }
        }

        private void OnReset(object sender, EventArgs e)
        {
            Control.Invoke((MethodInvoker) delegate
            {
                Control.IssueCount = 0;
                Control.IssueCountText = "0 issues";
                Control.InspectionResults.Clear();
            });
        }
    }
}
