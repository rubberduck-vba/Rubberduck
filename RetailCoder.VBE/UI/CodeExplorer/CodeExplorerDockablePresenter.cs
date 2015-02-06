using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Grammar;
using System;
using Antlr4.Runtime.Tree;
using Rubberduck.UI;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.UI.CodeExplorer
{
    [ComVisible(false)]
    public class CodeExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly IRubberduckParser _parser;
        private CodeExplorerWindow Control { get { return UserControl as CodeExplorerWindow; } }

        public CodeExplorerDockablePresenter(IRubberduckParser parser, VBE vbe, AddIn addIn)
            : base(vbe, addIn, new CodeExplorerWindow())
        {
            _parser = parser;
            Control.SolutionTree.Font = new Font(Control.SolutionTree.Font, FontStyle.Bold);
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
            //Control.SolutionTree.AfterExpand += TreeViewAfterExpandNode;
            //Control.SolutionTree.AfterCollapse += TreeViewAfterCollapseNode;
        }

        private void NavigateExplorerTreeNode(object sender, SyntaxTreeNodeClickEventArgs e)
        {
            //var instruction = e.Instruction;

            //var project = instruction.Line.ProjectName;
            //var component = instruction.Line.ComponentName;

            //var vbProject = VBE.VBProjects.Cast<VBProject>()
            //                   .FirstOrDefault(p => p.Name == project);

            //VBComponent vbComponent = null;
            //if (vbProject != null)
            //{
            //    vbComponent = vbProject.VBComponents.Cast<VBComponent>()
            //                           .FirstOrDefault(c => c.Name == component);
            //}

            //if (vbComponent == null)
            //{
            //    return;
            //}

            //var codePane = vbComponent.CodeModule.CodePane;
            //var selection = instruction.QualifiedSelection;

            //if (selection.StartLine != 0)
            //{
            //   codePane.SetSelection(selection);
            //}
        }

        private void RefreshExplorerTreeView()
        {
            Control.SolutionTree.Nodes.Clear();
            var projects = VBE.VBProjects.Cast<VBProject>().OrderBy(project => project.Name);
            foreach (var vbProject in projects)
            {
                AddProjectNode(_parser.Parse(vbProject));
            }
        }

        private void RefreshExplorerTreeView(object sender, System.EventArgs e)
        {
            RefreshExplorerTreeView();
        }

        private void AddProjectNode(IEnumerable<VbModuleParseResult> modules)
        {
            var treeView = Control.SolutionTree;
            // todo: [re-]implement

            var projectNode = new TreeNode();
            //projectNode.Text = node.Instruction.Line.ProjectName + new string(' ', 2);
            projectNode.Text = modules.First().QualifiedName.ProjectName;

            //projectNode.Tag = node.Instruction;
            projectNode.ImageKey = "ClosedFolder";
            treeView.BackColor = treeView.BackColor;

            var moduleNodes = new ConcurrentBag<TreeNode>();

            //foreach(var module in node.ChildNodes)
            //{
            foreach (var module in modules)
            {
                var moduleNode = new TreeNode(module.QualifiedName.ModuleName);
                moduleNode.NodeFont = new Font(treeView.Font, FontStyle.Regular);
                //todo: re-implement image & tag
                //    moduleNode.ImageKey = GetImageKeyForNode(module);
                //    moduleNode.SelectedImageKey = moduleNode.ImageKey;
                //    moduleNode.Tag = module.Instruction;

                for (var i = 0; i < module.ParseTree.ChildCount; i++)
                {
                    var child = module.ParseTree.GetChild(i);
                    if (child != null)
                    {
                        if (!string.IsNullOrWhiteSpace(child.GetText()))
                        {
                            //todo: find appropriate text from parse tree
                            moduleNode.Nodes.Add(child.GetText());
                        }
                    }
                }

                //todo:remove  old implementation
                //    foreach (var member in module.ChildNodes)
                //    {
                //        if (string.IsNullOrEmpty(member.Instruction.Value.Trim()))
                //        {
                //            // don't make a tree context for comments
                //            continue;
                //        }

                //        if (member.ChildNodes != null)
                //        {
                //            moduleNode.Nodes.Add(AddCodeBlockNode(member));
                //        }
                //    }
                moduleNodes.Add(moduleNode);
            }

            projectNode.Nodes.AddRange(moduleNodes.ToArray());
            treeView.Nodes.Add(projectNode);
        }

        //private TreeNode AddCodeBlockNode(SyntaxTreeNode node)
        //{
        //    var codeBlockNode = new TreeNode(GetNodeText(node));
        //    codeBlockNode.NodeFont = new Font(Control.SolutionTree.Font, FontStyle.Regular);
        //    codeBlockNode.ImageKey = GetImageKeyForNode(node);
        //    codeBlockNode.SelectedImageKey = codeBlockNode.ImageKey;
        //    codeBlockNode.Tag = node.Instruction;

        //    if (node.ChildNodes == null)
        //    {
        //        return codeBlockNode;
        //    }

        //    foreach (var member in node.ChildNodes)
        //    {
        //        if (string.IsNullOrEmpty(member.Instruction.Value.Trim()))
        //        {
        //            // don't make a tree context for comments
        //            continue;
        //        }

        //        var memberNode = new TreeNode(GetNodeText(member));

        //        memberNode.ToolTipText = string.Format("{0} (line {1})", 
        //                                        member.GetType().Name,
        //                                        member.Instruction.Line.StartLineNumber);

        //        memberNode.NodeFont = new Font(Control.SolutionTree.Font, FontStyle.Regular);
        //        memberNode.ImageKey = GetImageKeyForNode(member);
        //        memberNode.SelectedImageKey = memberNode.ImageKey;
        //        memberNode.Tag = member.Instruction;

        //        if (member.ChildNodes != null)
        //        {
        //            foreach (var child in member.ChildNodes)
        //            {
        //                memberNode.Nodes.Add(AddCodeBlockNode(child));
        //            }
        //        }

        //        codeBlockNode.Nodes.Add(memberNode);
        //    }

        //    return codeBlockNode;
        //}

        //private void TreeViewAfterExpandNode(object sender, TreeViewEventArgs e)
        //{
        //    if (!e.Node.ImageKey.Contains("Folder"))
        //    {
        //        return;
        //    }

        //    e.Node.ImageKey = "OpenFolder";
        //    e.Node.SelectedImageKey = e.Node.ImageKey;
        //}

        //private void TreeViewAfterCollapseNode(object sender, TreeViewEventArgs e)
        //{
        //    if (!e.Node.ImageKey.Contains("Folder"))
        //    {
        //        return;
        //    }

        //    e.Node.ImageKey = "ClosedFolder";
        //    e.Node.SelectedImageKey = e.Node.ImageKey;
        //}

        //private string GetImageKeyForNode(SyntaxTreeNode node)
        //{
        //    if (node is ModuleNode)
        //    {
        //        return (node as ModuleNode).IsClassModule
        //            ? (node.ChildNodes != null 
        //                && node.ChildNodes.OfType<ProcedureNode>().Any()
        //                && node.ChildNodes.OfType<ProcedureNode>().All(childNode => childNode.ChildNodes != null && !childNode.ChildNodes.Any()))
        //                ? "PublicInterface"
        //                : "ClassModule"
        //            : "StandardModule";
        //    }

        //    if (node is OptionNode)
        //    {
        //        return "Option";
        //    }

        //    if (node is ProcedureNode)
        //    {
        //        var propertyTypes = new[] {ProcedureKind.PropertyGet, ProcedureKind.PropertyLet, ProcedureKind.PropertySet};
        //        var procNode = (node as ProcedureNode);
        //        if (procNode.Accessibility == ReservedKeywords.Public)
        //        {
        //            return propertyTypes.Any(pt => pt == procNode.Kind) ? "PublicProperty" : "PublicMethod";
        //        }
        //        if (procNode.Accessibility == ReservedKeywords.Friend)
        //        {
        //            return propertyTypes.Any(pt => pt == procNode.Kind) ? "FriendProperty" : "FriendMethod";
        //        }
        //        if (procNode.Accessibility == ReservedKeywords.Private)
        //        {
        //            return propertyTypes.Any(pt => pt == procNode.Kind) ? "PrivateProperty" : "PrivateMethod";
        //        }
        //    }

        //    if (node is UserDefinedTypeNode)
        //    {
        //        var typeNode = (node as UserDefinedTypeNode);
        //        if (typeNode.Accessibility == ReservedKeywords.Public)
        //        {
        //            return "PublicType";
        //        }
        //        if (typeNode.Accessibility == ReservedKeywords.Friend)
        //        {
        //            return "FriendType";
        //        }
        //        if (typeNode.Accessibility == ReservedKeywords.Private)
        //        {
        //            return "PrivateType";
        //        }
        //    }

        //    if (node is EnumNode)
        //    {
        //        var typeNode = (node as EnumNode);
        //        if (typeNode.Accessibility == ReservedKeywords.Public)
        //        {
        //            return "PublicEnum";
        //        }
        //        if (typeNode.Accessibility == ReservedKeywords.Friend)
        //        {
        //            return "FriendEnum";
        //        }
        //        if (typeNode.Accessibility == ReservedKeywords.Private)
        //        {
        //            return "PrivateEnum";
        //        }
        //    }

        //    if (node is ConstDeclarationNode)
        //    {
        //        var accessbility = (node as DeclarationNode).Accessibility;
        //        if (accessbility == ReservedKeywords.Private)
        //        {
        //            return "PrivateConst";
        //        }
        //        if (accessbility == ReservedKeywords.Friend)
        //        {
        //            return "FriendConst";
        //        }

        //        return "PublicConst";
        //    }

        //    if (node is VariableDeclarationNode)
        //    {
        //        var accessbility = (node as DeclarationNode).Accessibility;
        //        if (accessbility == ReservedKeywords.Private)
        //        {
        //            return "PrivateField";
        //        }
        //        if (accessbility == ReservedKeywords.Friend)
        //        {
        //            return "FriendField";
        //        }

        //        return "PublicField";
        //    }

        //    if (node is CodeBlockNode)
        //    {
        //        return "CodeBlock";
        //    }

        //    if (node is IdentifierNode)
        //    {
        //        return "Identifier";
        //    }

        //    if (node is ParameterNode)
        //    {
        //        return "Parameter";
        //    }

        //    if (node is AssignmentNode)
        //    {
        //        return "Assignment";
        //    }

        //    if (node is UserDefinedTypeMemberNode)
        //    {
        //        return "PublicField";
        //    }

        //    if (node is EnumMemberNode)
        //    {
        //        return "EnumItem";
        //    }

        //    if (node is LabelNode)
        //    {
        //        return "Label";
        //    }

        //    return "Operation";
        //}

        //private string GetNodeText(SyntaxTreeNode node)
        //{
        //    if (node is ProcedureNode)
        //    {
        //        var procNode = node as ProcedureNode;
        //        var propertyTypes = new[] { ProcedureKind.PropertyGet, ProcedureKind.PropertyLet, ProcedureKind.PropertySet };
        //        if (propertyTypes.Any(pt => pt == procNode.Kind))
        //        {
        //            var kind = procNode.Kind == ProcedureKind.PropertyGet
        //                ? ReservedKeywords.Get
        //                : procNode.Kind == ProcedureKind.PropertyLet
        //                    ? ReservedKeywords.Let
        //                    : ReservedKeywords.Set;

        //            return string.Format("{0} ({1})", procNode.Identifier.Name, kind);
        //        }
        //        return procNode.Identifier.Name;
        //    }

        //    if (node is UserDefinedTypeNode)
        //    {
        //        return ((UserDefinedTypeNode) node).Identifier.Name;
        //    }

        //    if (node is EnumNode)
        //    {
        //        return ((EnumNode) node).Identifier.Name;
        //    }

        //    if (node is IdentifierNode)
        //    {
        //        return ((IdentifierNode) node).Name;
        //    }

        //    return node.Instruction.Value.Trim();
        //}
    }
}
