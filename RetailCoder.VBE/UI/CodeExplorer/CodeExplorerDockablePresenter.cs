using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;
using AddIn = Microsoft.Vbe.Interop.AddIn;
using Font = System.Drawing.Font;
using Selection = Rubberduck.Extensions.Selection;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly IRubberduckParser _parser;
        private CodeExplorerWindow Control { get { return UserControl as CodeExplorerWindow; } }

        public CodeExplorerDockablePresenter(IRubberduckParser parser, VBE vbe, AddIn addIn, CodeExplorerWindow view)
            : base(vbe, addIn, view)
        {
            _parser = parser;
            RegisterControlEvents();
            RefreshExplorerTreeView();
            Control.SolutionTree.Refresh();
        }

        private void RegisterControlEvents()
        {
            if (Control == null)
            {
                return;
            }

            Control.RefreshTreeView += RefreshExplorerTreeView;
            Control.NavigateTreeNode += NavigateExplorerTreeNode;
            Control.SolutionTree.AfterExpand += TreeViewAfterExpandNode;
            Control.SolutionTree.AfterCollapse += TreeViewAfterCollapseNode;
        }

        private void NavigateExplorerTreeNode(object sender, TreeNodeNavigateCodeEventArgs e)
        {
            if (!e.Node.IsExpanded)
            {
                e.Node.Expand();
            }

            if (e.Selection.StartLine != 0)
            {
                //hack: get around issue where a node's selection seems to ignore a procedure's (or enum's) signature
                var selection = new Selection(e.Selection.StartLine,
                                                1,
                                                e.Selection.EndLine,
                                                e.Selection.EndColumn == 1 ? 0 : e.Selection.EndColumn //fixes off by one error when navigating the module
                                              );
                VBE.SetSelection(new QualifiedSelection(e.QualifiedName, selection));
            }
        }

        private void RefreshExplorerTreeView()
        {
            Control.SolutionTree.Nodes.Clear();

            var projects = VBE.VBProjects.Cast<VBProject>();
            foreach (var vbProject in projects)
            {
                var project = vbProject;
                Task.Run(() =>
                {
                    var node = new TreeNode(project.Name + " (parsing...)");
                    node.ImageKey = "Hourglass";
                    node.SelectedImageKey = node.ImageKey;

                    Control.Invoke((MethodInvoker) delegate
                    {
                        Control.SolutionTree.Nodes.Add(node);
                        AddProjectNodes(project, node);
                    });
                });
            }

            Control.SolutionTree.BackColor = Control.SolutionTree.BackColor;
        }

        private void RefreshExplorerTreeView(object sender, System.EventArgs e)
        {
            Control.Cursor = Cursors.WaitCursor;
            RefreshExplorerTreeView();
            Control.Cursor = Cursors.Default;
        }

        private void AddProjectNodes(VBProject project, TreeNode root)
        {
            var treeView = Control.SolutionTree;
            Control.Invoke((MethodInvoker) async delegate
            {
                await AddModuleNodesAsync(project, treeView.Font, root);
                root.Text = project.Name;
                root.Tag = new QualifiedSelection();
                root.ImageKey = "ClosedFolder";
            });
        }

        private async Task AddModuleNodesAsync(VBProject project, Font font, TreeNode root)
        {
            foreach (VBComponent vbComponent in project.VBComponents)
            {
                var component = vbComponent;
                    var qualifiedName = component.QualifiedName();
                    var node = new TreeNode(component.Name + " (parsing...)");
                    node.ImageKey = "Hourglass";
                    node.SelectedImageKey = node.ImageKey;
                    node.NodeFont = new Font(font, FontStyle.Regular);

                    root.Nodes.Add(node);

                    var getModuleNode = new Task<TreeNode[]>(() => ParseModule(component, ref qualifiedName));
                    getModuleNode.Start();
                    node.Nodes.AddRange(getModuleNode.Result);

                    node.Text = component.Name;
                    node.ImageKey = "StandardModule";
                    node.SelectedImageKey = node.ImageKey;
            }
        }

        private TreeNode[] ParseModule(VBComponent component, ref QualifiedModuleName qualifiedName)
        {
            var moduleNode = _parser.Parse(component).ParseTree.GetContexts<TreeViewListener, TreeNode>(new TreeViewListener(qualifiedName)).Single();
            return moduleNode.Context.Nodes.Cast<TreeNode>().ToArray();
        }

        private void TreeViewAfterExpandNode(object sender, TreeViewEventArgs e)
        {
            if (!e.Node.ImageKey.Contains("Folder"))
            {
                return;
            }

            e.Node.ImageKey = "OpenFolder";
            e.Node.SelectedImageKey = e.Node.ImageKey;
        }

        private void TreeViewAfterCollapseNode(object sender, TreeViewEventArgs e)
        {
            if (!e.Node.ImageKey.Contains("Folder"))
            {
                return;
            }

            e.Node.ImageKey = "ClosedFolder";
            e.Node.SelectedImageKey = e.Node.ImageKey;
        }
    }
}
