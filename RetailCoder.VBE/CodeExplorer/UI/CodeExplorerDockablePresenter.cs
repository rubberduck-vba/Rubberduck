using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;

namespace Rubberduck.CodeExplorer.UI
{
    [ComVisible(false)]
    public class CodeExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly Parser _parser;
        private readonly CodeExplorerWindow _control;
        private readonly Window _window;

        public CodeExplorerDockablePresenter(Parser parser, VBE vbe, AddIn addIn)
            : base(vbe, addIn, "Code Explorer", new CodeExplorerWindow())
        {
            _parser = parser;
            _control = base.UserControl as CodeExplorerWindow;
            if (_control != null)
            {
                _control.RefreshTreeView += RefreshExplorerTreeView;
                _control.NavigateTreeNode += NavigateExplorerTreeNode;
            }
        }

        private void NavigateExplorerTreeNode(object sender, SyntaxTreeNodeClickEventArgs e)
        {
            var instruction = e.Node.Instruction;
            var selection = new Selection(instruction.Line.EndLineNumber, instruction.StartColumn, instruction.Line.EndLineNumber, instruction.EndColumn);

            var project = instruction.Line.ProjectName;
            var component = instruction.Line.ComponentName;

            var codeModule = VBE.VBProjects.Cast<VBProject>()
                                .First(p => p.Name == project)
                                .VBComponents.Cast<VBComponent>()
                                .First(c => c.Name == component)
                                .CodeModule;

            codeModule.CodePane.SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
            codeModule.CodePane.Show();
        }

        private void RefreshExplorerTreeView(object sender, System.EventArgs e)
        {
            _control.SolutionTree.Nodes.Clear();
            var projects = VBE.VBProjects.Cast<VBProject>();
            foreach (var project in projects)
            {
                AddProjectNode(_parser.Parse(project));
            }
        }

        private void AddProjectNode(SyntaxTreeNode node)
        {
            var tree = _control.SolutionTree;
            var projectNode = new TreeNode(node.Scope);
        }
    }
}
