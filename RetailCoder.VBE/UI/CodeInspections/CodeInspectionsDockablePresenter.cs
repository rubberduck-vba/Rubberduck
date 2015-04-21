using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private IList<ICodeInspectionResult> _results;
        private readonly IInspector _inspector;

        public CodeInspectionsDockablePresenter(IInspector inspector, VBE vbe, AddIn addin, CodeInspectionsWindow window)
            :base(vbe, addin, window)
        {
            _inspector = inspector;
            _inspector.IssuesFound += OnIssuesFound;
            _inspector.Reset += OnReset;

            Control.RefreshCodeInspections += OnRefreshCodeInspections;
            Control.NavigateCodeIssue += OnNavigateCodeIssue;
            Control.QuickFix += OnQuickFix;
            Control.CopyResults += OnCopyResultsToClipboard;
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
            var newCount = Control.IssueCount + e.Issues.Count;
            Control.IssueCount = newCount;
            Control.IssueCountText = string.Format("{0} issue" + (newCount != 1 ? "s" : string.Empty), newCount);
        }

        private void OnQuickFix(object sender, QuickFixEventArgs e)
        {
            e.QuickFix(VBE);
            OnRefreshCodeInspections(null, EventArgs.Empty);
        }

        public override void Show()
        {
            base.Show();
            Refresh();
        }

        private void OnNavigateCodeIssue(object sender, NavigateCodeEventArgs e)
        {
            try
            {
                var location = VBE.FindInstruction(e.QualifiedName, e.Selection);
                location.CodeModule.CodePane.SetSelection(e.Selection);

                var codePane = location.CodeModule.CodePane;
                var selection = location.Selection;
                codePane.SetSelection(selection);
            }
            catch (Exception exception)
            {
                Debug.Assert(false, exception.ToString());
            }
        }

        private void OnRefreshCodeInspections(object sender, EventArgs e)
        {
            Refresh();
        }

        private async void Refresh()
        {
            Control.Cursor = Cursors.WaitCursor;

            try
            {
                if (VBE != null)
                {
                    _results = await _inspector.FindIssuesAsync(VBE.ActiveVBProject);
                    Control.SetContent(_results.Select(item => new CodeInspectionResultGridViewItem(item))
                        .OrderBy(item => item.Component)
                        .ThenBy(item => item.Line));

                    if (!_results.Any())
                    {
                        Control.QuickFixButton.Enabled = false;
                    }
                }
            }
            catch (COMException)
            {
                // swallow
            }
            finally
            {
                Control.Cursor = Cursors.Default;
            }
        }

        private void OnReset(object sender, EventArgs e)
        {
            Control.IssueCount = 0;
            Control.IssueCountText = "0 issues";
            Control.InspectionResults.Clear();
        }
    }
}
