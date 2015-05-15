using System;
using System.Collections.Generic;
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

        // indicates that the _parseResults are no longer in sync with the UI
        private bool _needsResync;

        private void OnParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            ToggleParsingStatus(false);
            if (sender == this)
            {
                _needsResync = false;
                _parseResults = e.ParseResults;
                Task.Run(() => RefreshAsync());
            }
            else
            {
                _parseResults = e.ParseResults;
                _needsResync = true;
            }
        }

        private void OnParsing(object sender, EventArgs e)
        {
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

        private int _issues;
        private void OnIssuesFound(object sender, InspectorIssuesFoundEventArg e)
        {
            Interlocked.Add(ref _issues, e.Issues.Count);
            Control.Invoke((MethodInvoker) delegate
            {
                var newCount = _issues;
                Control.SetIssuesStatus(newCount);
            });
        }

        private void OnQuickFix(object sender, QuickFixEventArgs e)
        {
            e.QuickFix(VBE);
            _needsResync = true;
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
            Task.Run(() => RefreshAsync()).ContinueWith(t =>
            {
                Control.SetIssuesStatus(_results.Count, true);
            });
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
                    if (_parseResults == null || _needsResync)
                    {
                        _inspector.Parse(VBE, this);
                        return;
                    }

                    var parseResults = _parseResults.SingleOrDefault(p => p.Project == VBE.ActiveVBProject);
                    if (parseResults == null || _needsResync)
                    {
                        _inspector.Parse(VBE, this);
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
            catch (COMException exception)
            {
                // swallow
            }
            finally
            {
                Control.Invoke((MethodInvoker) delegate
                {
                    Control.Cursor = Cursors.Default;
                    Control.SetIssuesStatus(_issues, true);
                    Control.EnableRefresh();
                });
            }
        }

        private void OnReset(object sender, EventArgs e)
        {
            _issues = 0;
            Control.Invoke((MethodInvoker) delegate
            {
                Control.SetIssuesStatus(_issues);
                Control.InspectionResults.Clear();
            });
        }
    }
}
