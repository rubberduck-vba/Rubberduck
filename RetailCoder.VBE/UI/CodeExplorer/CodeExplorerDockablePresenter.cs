using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.UI.CodeExplorer
{
    [ComVisible(false)]
    public class CodeExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly Parser _parser;
        private readonly CodeExplorerWindow _control;

        public CodeExplorerDockablePresenter(Parser parser, VBE vbe, AddIn addIn)
            : base(vbe, addIn, "Code Explorer", new CodeExplorerWindow())
        {
            _parser = parser;
            _control = base.UserControl as CodeExplorerWindow;
            if (_control != null)
            {
                _control.RefreshTreeView += RefreshExplorerTreeView;
                _control.NavigateTreeNode += NavigateExplorerTreeNode;
                _control.SolutionTree.AfterExpand += TreeViewAfterExpandNode;
                _control.SolutionTree.AfterCollapse += TreeViewAfterCollapseNode;
            }
        }

        private void NavigateExplorerTreeNode(object sender, SyntaxTreeNodeClickEventArgs e)
        {
            var instruction = e.Instruction;
            var selection = new Selection(instruction.Line.EndLineNumber, instruction.StartColumn, instruction.Line.EndLineNumber, instruction.EndColumn);

            var project = instruction.Line.ProjectName;
            var component = instruction.Line.ComponentName;

            var vbProject = VBE.VBProjects.Cast<VBProject>()
                               .FirstOrDefault(p => p.Name == project);

            VBComponent vbComponent = null;
            if (vbProject != null)
            {
                vbComponent = vbProject.VBComponents.Cast<VBComponent>()
                                       .FirstOrDefault(c => c.Name == component);
            }

            if (vbComponent == null)
            {
                return;
            }

            if (selection.StartLine != 0)
            {
                vbComponent.CodeModule.CodePane.SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
            }

            vbComponent.CodeModule.CodePane.Show();
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
            var treeView = _control.SolutionTree;
            var projectNode = new TreeNode();
            projectNode.Text = node.Instruction.Line.ProjectName;
            projectNode.Tag = node.Instruction;
            projectNode.ImageKey = "ClosedFolder";

            foreach (var module in node.ChildNodes)
            {
                var moduleNode = new TreeNode(((ModuleNode) module).Identifier.Name);
                moduleNode.ImageKey = GetImageKeyForNode(module);
                moduleNode.SelectedImageKey = moduleNode.ImageKey;
                moduleNode.Tag = module.Instruction;

                foreach (var member in module.ChildNodes)
                {
                    var memberNode = new TreeNode(GetNodeText(member));
                    memberNode.ToolTipText = string.Format("Line {0}", member.Instruction.Line.StartLineNumber);
                    memberNode.ImageKey = GetImageKeyForNode(member);
                    memberNode.SelectedImageKey = memberNode.ImageKey;
                    memberNode.Tag = member.Instruction;
                    moduleNode.Nodes.Add(memberNode);
                }

                projectNode.Nodes.Add(moduleNode);
            }            

            treeView.Nodes.Add(projectNode);
        }

        private void TreeViewAfterExpandNode(object sender, TreeViewEventArgs e)
        {
            if (!e.Node.ImageKey.Contains("Folder"))
            {
                return;
            }

            e.Node.ImageKey = "OpenedFolder";
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

        private string GetImageKeyForNode(SyntaxTreeNode node)
        {
            if (node is ModuleNode)
            {
                return (node as ModuleNode).IsClassModule
                    ? "ClassModule"
                    : "StandardModule";
            }

            if (node is OptionNode)
            {
                return "Option";
            }

            if (node is ProcedureNode)
            {
                var procNode = (node as ProcedureNode);
                if (procNode.Accessibility == ReservedKeywords.Public)
                {
                    return procNode.Kind == ProcedureKind.Property ? "PublicProperty" : "PublicMethod";
                }
                if (procNode.Accessibility == ReservedKeywords.Friend)
                {
                    return procNode.Kind == ProcedureKind.Property ? "FriendProperty" : "FriendMethod";
                }
                if (procNode.Accessibility == ReservedKeywords.Private)
                {
                    return procNode.Kind == ProcedureKind.Property ? "PrivateProperty" : "PrivateMethod";
                }
            }

            if (node is UserDefinedTypeNode)
            {
                var typeNode = (node as UserDefinedTypeNode);
                if (typeNode.Accessibility == ReservedKeywords.Public)
                {
                    return "PublicType";
                }
                if (typeNode.Accessibility == ReservedKeywords.Friend)
                {
                    return "FriendType";
                }
                if (typeNode.Accessibility == ReservedKeywords.Private)
                {
                    return "PrivateType";
                }
            }

            if (node is EnumNode)
            {
                var typeNode = (node as EnumNode);
                if (typeNode.Accessibility == ReservedKeywords.Public)
                {
                    return "PublicEnum";
                }
                if (typeNode.Accessibility == ReservedKeywords.Friend)
                {
                    return "FriendEnum";
                }
                if (typeNode.Accessibility == ReservedKeywords.Private)
                {
                    return "PrivateEnum";
                }
            }

            if (node is ConstDeclarationNode)
            {
                var accessbility = (node as DeclarationNode).Accessibility;
                if (accessbility == ReservedKeywords.Private)
                {
                    return "PrivateConst";
                }
                if (accessbility == ReservedKeywords.Friend)
                {
                    return "FriendConst";
                }

                return "PublicConst";
            }

            if (node is VariableDeclarationNode)
            {
                var accessbility = (node as DeclarationNode).Accessibility;
                if (accessbility == ReservedKeywords.Private)
                {
                    return "PrivateField";
                }
                if (accessbility == ReservedKeywords.Friend)
                {
                    return "FriendField";
                }

                return "PublicField";
            }

            return "ClassModule"; // todo: find a better default.
        }

        private string GetNodeText(SyntaxTreeNode node)
        {
            if (node is ProcedureNode)
            {
                return ((ProcedureNode) node).Identifier.Name;
            }

            if (node is OptionNode)
            {
                return node.Instruction.Value;
            }

            if (node is UserDefinedTypeNode)
            {
                return ((UserDefinedTypeNode) node).Identifier.Name;
            }

            if (node is EnumNode)
            {
                return ((EnumNode) node).Identifier.Name;
            }

            if (node is DeclarationNode)
            {
                if (node.ChildNodes.Count() == 1)
                {
                    return ((IdentifierNode) node.ChildNodes.First()).Name;
                }
                else
                {
                    return node.Instruction.Value;
                }
            }

            return node.Instruction.Value;
        }
    }
}
