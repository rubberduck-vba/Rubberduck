using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA;
using Rubberduck.VBA.ParseTreeListeners;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.UnitTesting;
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
            Control.AddComponent += AddComponent;
            Control.AddTestModule += AddTestModule;
            Control.ToggleFolders += ToggleFolders;
            Control.ShowDesigner += ShowDesigner;
            Control.DisplayStyleChanged += DisplayStyleChanged;
            Control.RunAllTests += ContextMenuRunAllTests;
        }

        public event EventHandler RunAllTests;
        private void ContextMenuRunAllTests(object sender, EventArgs e)
        {
            var handler = RunAllTests;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void DisplayStyleChanged(object sender, EventArgs e)
        {
            RefreshExplorerTreeView();
        }

        private void AddTestModule(object sender, EventArgs e)
        {
            NewUnitTestModuleCommand.NewUnitTestModule(VBE);
            RefreshExplorerTreeView();
        }

        private void ShowDesigner(object sender, System.EventArgs e)
        {
            var node = Control.SolutionTree.SelectedNode;
            if (node != null && node.Tag != null)
            {
                var selection = (QualifiedSelection)node.Tag;
                var module = VBE.FindCodeModules(selection.QualifiedName).FirstOrDefault();
                if (module == null)
                {
                    return;
                }

                try
                {
                    module.Parent.DesignerWindow().Visible = true;
                }
                catch
                {
                    Control.ShowDesignerButton.Enabled = false;
                }
            }
        }

        private bool _showFolders = true;
        private void ToggleFolders(object sender, System.EventArgs e)
        {
            _showFolders = !_showFolders;
            RefreshExplorerTreeView();
        }

        private void AddComponent(object sender, AddComponentEventArgs e)
        {
            var project = VBE.ActiveVBProject;
            if (project == null)
            {
                return;
            }

            project.VBComponents.Add(e.ComponentType);
            RefreshExplorerTreeView();
        }

        private void NavigateExplorerTreeNode(object sender, TreeNodeNavigateCodeEventArgs e)
        {
            if (e.Selection.StartLine != 0)
            {
                //hack: get around issue where a node's selection seems to ignore a procedure's (or enum's) signature
                //todo: determiner if this "temp fix" is still needed.
                var selection = new Selection(e.Selection.StartLine,
                                                1,
                                                e.Selection.EndLine,
                                                e.Selection.EndColumn == 1 ? 0 : e.Selection.EndColumn //fixes off by one error when navigating the module
                                              );
                VBE.SetSelection(new QualifiedSelection(e.QualifiedName, selection));
            }
        }

        private void RefreshExplorerTreeView(object sender, System.EventArgs e)
        {
            RefreshExplorerTreeView();
        }

        private async void RefreshExplorerTreeView()
        {
            Control.SolutionTree.Nodes.Clear();
            Control.ShowDesignerButton.Enabled = false;

            var projects = VBE.VBProjects.Cast<VBProject>();
            foreach (var vbProject in projects)
            {
                var project = vbProject;
                await Task.Run(() =>
                {
                    var node = new TreeNode(project.Name + " (parsing...)");
                    node.ImageKey = "Hourglass";
                    node.SelectedImageKey = node.ImageKey;

                    Control.Invoke((MethodInvoker)delegate
                    {
                        Control.SolutionTree.Nodes.Add(node);
                        AddProjectNodes(project, node);
                    });
                });
            }

            Control.SolutionTree.BackColor = Control.SolutionTree.BackColor;
        }

        private void AddProjectNodes(VBProject project, TreeNode root)
        {
            var treeView = Control.SolutionTree;
            Control.Invoke((MethodInvoker)async delegate
            {
                root.Text = project.Name;
                if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
                {
                    root.ImageKey = "Locked";
                }
                else
                {
                    root.ImageKey = "ClosedFolder";
                    var nodes = (await CreateModuleNodesAsync(project, treeView.Font)).ToArray();
                    AddProjectFolders(project, root, nodes);
                    root.Expand();
                }
            });
        }

        private static readonly IDictionary<vbext_ComponentType, string> ComponentTypeIcons =
            new Dictionary<vbext_ComponentType, string>
            {
                { vbext_ComponentType.vbext_ct_StdModule, "StandardModule"},
                { vbext_ComponentType.vbext_ct_ClassModule, "ClassModule"},
                { vbext_ComponentType.vbext_ct_Document, "OfficeDocument"},
                { vbext_ComponentType.vbext_ct_ActiveXDesigner, "ClassModule"},
                { vbext_ComponentType.vbext_ct_MSForm, "Form"}
            };

        private void AddProjectFolders(VBProject project, TreeNode root, TreeNode[] components)
        {
            var documentNodes = components.Where(node => node.ImageKey == "OfficeDocument")
                                          .OrderBy(node => node.Text)
                                          .ToArray();
            if (project.VBComponents.Cast<VBComponent>()
                       .Any(component => component.Type == vbext_ComponentType.vbext_ct_Document))
            {
                AddFolderNode(root, "Documents", "ClosedFolder", documentNodes);
            }

            var formsNodes = components.Where(node => node.ImageKey == "Form")
                                       .OrderBy(node => node.Text)
                                       .ToArray();
            if (project.VBComponents.Cast<VBComponent>()
                       .Any(component => component.Type == vbext_ComponentType.vbext_ct_MSForm))
            {
                AddFolderNode(root, "Forms", "ClosedFolder", formsNodes);
            }

            var stdModulesNodes = components.Where(node => node.ImageKey == "StandardModule")
                                            .OrderBy(node => node.Text)
                                            .ToArray();
            if (project.VBComponents.Cast<VBComponent>()
                       .Any(component => component.Type == vbext_ComponentType.vbext_ct_StdModule))
            {
                AddFolderNode(root, "Standard Modules", "ClosedFolder", stdModulesNodes);
            }

            var classModulesNodes = components.Where(node => node.ImageKey == "ClassModule")
                                              .OrderBy(node => node.Text)
                                              .ToArray();
            if (project.VBComponents.Cast<VBComponent>()
                       .Any(component => component.Type == vbext_ComponentType.vbext_ct_ClassModule
                                      || component.Type == vbext_ComponentType.vbext_ct_ActiveXDesigner))
            {
                AddFolderNode(root, "Class Modules", "ClosedFolder", classModulesNodes);
            }
        }

        private void AddFolderNode(TreeNode root, string text, string imageKey, TreeNode[] nodes)
        {
            if (_showFolders)
            {
                var node = root.Nodes.Add(text);
                node.ImageKey = imageKey;
                node.SelectedImageKey = imageKey;
                node.Nodes.AddRange(nodes);
                node.Expand();
            }
            else
            {
                root.Nodes.AddRange(nodes);
            }
        }

        private async Task<IEnumerable<TreeNode>> CreateModuleNodesAsync(VBProject project, Font font)
        {
            var result = new List<TreeNode>();
            foreach (VBComponent vbComponent in project.VBComponents)
            {
                var component = vbComponent;
                var qualifiedName = component.QualifiedName();

                var members = await Task.Run(() => ParseModule(component, ref qualifiedName));

                var node = members.Context;
                node.ImageKey = ComponentTypeIcons[component.Type];
                node.SelectedImageKey = node.ImageKey;
                //node.NodeFont = new Font(font, FontStyle.Regular);
                //node.Text = component.Name;

                result.Add(node);
            }

            return result;
        }

        private QualifiedContext<TreeNode> ParseModule(VBComponent component, ref QualifiedModuleName qualifiedName)
        {
            return _parser.Parse(component).ParseTree.GetContexts<TreeViewListener, TreeNode>(new TreeViewListener(qualifiedName, Control.DisplayStyle)).Single();
        }
    }
}
