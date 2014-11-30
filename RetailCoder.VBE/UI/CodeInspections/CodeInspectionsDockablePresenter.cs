using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class CodeInspectionsDockablePresenter : DockablePresenterBase
    {
        private readonly Parser _parser;
        private CodeInspectionsWindow Control { get { return UserControl as CodeInspectionsWindow; } }

        private readonly IList<IInspection> _inspections;

        public CodeInspectionsDockablePresenter(Parser parser, IEnumerable<IInspection> inspections, VBE vbe, AddIn addin) 
            : base(vbe, addin, new CodeInspectionsWindow())
        {
            _parser = parser;
            _inspections = inspections.ToList();

            Control.RefreshCodeInspections += OnRefreshCodeInspections;
        }
        
        private void OnRefreshCodeInspections(object sender, EventArgs e)
        {
            var code = _parser.Parse(VBE.ActiveVBProject);
            var results = new List<CodeInspectionResultBase>();
            foreach (var inspection in _inspections.Where(inspection => inspection.IsEnabled))
            {
                var result = inspection.Inspect(code).ToArray();
                if (result.Length != 0)
                {
                    results.AddRange(result);
                }
            }

            DrawResultTree(results);
        }

        private void DrawResultTree(IEnumerable<CodeInspectionResultBase> results)
        {
            var tree = Control.CodeInspectionResultsTree;
            tree.Nodes.Clear();

            foreach (var result in results.OrderBy(r => r.Severity))
            {
                var node = new TreeNode(result.Message);

                tree.Nodes.Add(result.Message);
            }
        }
    }
}
