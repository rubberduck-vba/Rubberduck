using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Listeners;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UnitTesting;
using Rubberduck.VBA;

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
            Control.RunInspections += ContextMenuRunInspections;
        }

        public event EventHandler RunInspections;
        private void ContextMenuRunInspections(object sender, EventArgs e)
        {
            var handler = RunInspections;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
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

        private void ShowDesigner(object sender, EventArgs e)
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
        private void ToggleFolders(object sender, EventArgs e)
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

        private void RefreshExplorerTreeView(object sender, EventArgs e)
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
                        Control.SolutionTree.Refresh();
                        AddProjectNodes(project, node);
                    });
                });
            }

            // note: is this really needed?
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
                    var nodes = (await CreateModuleNodesAsync(project)).ToArray();
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

        private async Task<IEnumerable<TreeNode>> CreateModuleNodesAsync(VBProject project)
        {
            var result = new List<TreeNode>();
            var parseResult = _parser.Parse(project);
            foreach (var componentParseResult in parseResult.ComponentParseResults)
            {
                var component = componentParseResult.Component;
                var members = parseResult.Declarations.Items
                    .Where(declaration => declaration.ParentScope == component.Collection.Parent.Name + "." + component.Name)
                    .OrderBy(declaration => declaration.IdentifierName);

                var node = new TreeNode(component.Name);
                node.ImageKey = ComponentTypeIcons[component.Type];
                node.SelectedImageKey = node.ImageKey;
                node.Tag = new QualifiedSelection(componentParseResult.QualifiedName, componentParseResult.Context.GetSelection());

                foreach (var declaration in members)
                {
                    var text = GetNodeText(declaration);
                    var child = new TreeNode(text);
                    child.ImageKey = GetImageKeyForDeclaration(declaration);
                    child.SelectedImageKey = child.ImageKey;
                    child.Tag = new QualifiedSelection(declaration.QualifiedName.ModuleScope, declaration.Selection);

                    if (declaration.DeclarationType == DeclarationType.UserDefinedType
                        || declaration.DeclarationType == DeclarationType.Enumeration)
                    {
                        var subDeclaration = declaration;
                        var subMembers = parseResult.Declarations.Items.Where(item => item.ParentScope == subDeclaration.Scope + "." + subDeclaration.IdentifierName);
                        foreach (var subMember in subMembers)
                        {
                            var subChild = new TreeNode(subMember.IdentifierName);
                            subChild.ImageKey = GetImageKeyForDeclaration(subMember);
                            subChild.SelectedImageKey = subChild.ImageKey;
                            subChild.Tag = new QualifiedSelection(subMember.QualifiedName.ModuleScope, subMember.Selection);
                            child.Nodes.Add(subChild);
                        }
                    }

                    node.Nodes.Add(child);
                }

                result.Add(node);
            }

            return result;
        }

        private string GetNodeText(Declaration declaration)
        {
            if (Control.DisplayStyle == TreeViewDisplayStyle.MemberNames)
            {
                var result = declaration.IdentifierName;
                if (declaration.DeclarationType == DeclarationType.PropertyGet)
                {
                    result += " (" + Tokens.Get + ")";
                }
                else if (declaration.DeclarationType == DeclarationType.PropertyLet)
                {
                    result += " (" + Tokens.Let + ")";
                }
                else if (declaration.DeclarationType == DeclarationType.PropertySet)
                {
                    result += " (" + Tokens.Set + ")";
                }

                return result;
            }

            if (declaration.DeclarationType == DeclarationType.Procedure)
            {
                return ((VBAParser.SubStmtContext) declaration.Context).Signature();
            }

            if (declaration.DeclarationType == DeclarationType.Function)
            {
                return ((VBAParser.FunctionStmtContext)declaration.Context).Signature();
            }

            if (declaration.DeclarationType == DeclarationType.PropertyGet)
            {
                return ((VBAParser.PropertyGetStmtContext)declaration.Context).Signature();
            }

            if (declaration.DeclarationType == DeclarationType.PropertyLet)
            {
                return ((VBAParser.PropertyLetStmtContext)declaration.Context).Signature();
            }

            if (declaration.DeclarationType == DeclarationType.PropertySet)
            {
                return ((VBAParser.PropertySetStmtContext)declaration.Context).Signature();
            }

            return declaration.IdentifierName;
        }

        private string GetImageKeyForDeclaration(Declaration declaration)
        {
            var result = string.Empty;
            switch (declaration.DeclarationType)
            {
                case DeclarationType.Module:
                    break;
                case DeclarationType.Class:
                    break;
                case DeclarationType.Procedure:
                case DeclarationType.Function:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateMethod";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendMethod";
                        break;
                    }
                    result = "PublicMethod";
                    break;

                case DeclarationType.PropertyGet:
                case DeclarationType.PropertyLet:
                case DeclarationType.PropertySet:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateProperty";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendProperty";
                        break;
                    }
                    result = "PublicProperty";
                    break;

                case DeclarationType.Parameter:
                    break;
                case DeclarationType.Variable:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateField";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendField";
                        break;
                    }
                    result = "PublicField";
                    break;

                case DeclarationType.Constant:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateConst";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendConst";
                        break;
                    }
                    result = "PublicConst";
                    break;

                case DeclarationType.Enumeration:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateEnum";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendEnum";
                        break;
                    }
                    result = "PublicEnum";
                    break;

                case DeclarationType.EnumerationMember:
                    result = "EnumItem";
                    break;

                case DeclarationType.Event:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateEvent";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendEvent";
                        break;
                    }
                    result = "PublicEvent";
                    break;

                case DeclarationType.UserDefinedType:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        result = "PrivateType";
                        break;
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        result = "FriendType";
                        break;
                    }
                    result = "PublicType";
                    break;

                case DeclarationType.UserDefinedTypeMember:
                    result = "PublicField";
                    break;
                
                case DeclarationType.LibraryFunction:
                    result = "Identifier";
                    break;

                default:
                    throw new ArgumentOutOfRangeException();
            }

            return result;
        }
    }
}
