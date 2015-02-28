using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private IList<ICodeInspectionResult> _results;
        private IInspector _inspector;

        public CodeInspectionsDockablePresenter(IInspector inspector, VBE vbe, AddIn addin, CodeInspectionsWindow window)
            :base(vbe, addin, window)
        {
            _inspector = inspector;
            Control.RefreshCodeInspections += OnRefreshCodeInspections;
            Control.NavigateCodeIssue += OnNavigateCodeIssue;
            Control.QuickFix += OnQuickFix;
        }

        private void OnQuickFix(object sender, QuickFixEventArgs e)
        {
            e.QuickFix(VBE);
            OnRefreshCodeInspections(null, EventArgs.Empty);
        }

        public override void Show()
        {
            base.Show();

            if (VBE.ActiveVBProject != null)
            {
                OnRefreshCodeInspections(this, EventArgs.Empty);
            }
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
                System.Diagnostics.Debug.Assert(false, exception.ToString());
            }
        }

        private void OnRefreshCodeInspections(object sender, EventArgs e)
        {
            Control.Cursor = Cursors.WaitCursor;
            Refresh();
            Control.Cursor = Cursors.Default;
        }

        private void Refresh()
        {
            _results = this._inspector.FindIssues(VBE.ActiveVBProject);
            Control.SetContent(_results.Select(item => new CodeInspectionResultGridViewItem(item)).OrderBy(item => item.Component).ThenBy(item => item.Line));
        }
    }
}
