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

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private readonly IRubberduckParser _parser;
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private readonly IList<IInspection> _inspections;
        private List<CodeInspectionResultBase> _results;

        public CodeInspectionsDockablePresenter(IRubberduckParser parser, IEnumerable<IInspection> inspections, VBE vbe, AddIn addin, CodeInspectionsWindow window) 
            : base(vbe, addin, window)
        {
            _parser = parser;
            _inspections = inspections.ToList();

            Control.RefreshCodeInspections += OnRefreshCodeInspections;
            Control.NavigateCodeIssue += OnNavigateCodeIssue;
            Control.QuickFix += OnQuickFix;
        }

        private void OnQuickFix(object sender, QuickFixEventArgs e)
        {
            e.QuickFix(VBE);
            OnRefreshCodeInspections(null, EventArgs.Empty);
            //Control.FindNextIssue(); // note: decide if this is annoying or surprising, UX-wise
        }

        public override void Show()
        {
            if (VBE.ActiveVBProject != null)
            {
                OnRefreshCodeInspections(this, EventArgs.Empty);
            }
            base.Show();
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
            RefreshAsync();
            Control.Cursor = Cursors.Default;
        }

        private async void RefreshAsync()
        {
            var code = (_parser.Parse(VBE.ActiveVBProject)).ToList();

            var results = new ConcurrentBag<CodeInspectionResultBase>();
            var inspections = _inspections.Where(inspection => inspection.Severity != CodeInspectionSeverity.DoNotShow);
            Parallel.ForEach(inspections, inspection =>
            {
                var result = inspection.GetInspectionResults(code);
                foreach (var inspectionResult in result)
                {
                    results.Add(inspectionResult);
                }
            });

            _results = results.ToList();
            Control.SetContent(_results.Select(item => new CodeInspectionResultGridViewItem(item)).OrderBy(item => item.Component).ThenBy(item => item.Line));
        }
    }
}
