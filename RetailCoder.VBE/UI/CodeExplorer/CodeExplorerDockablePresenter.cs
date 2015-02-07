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
using Rubberduck.Extensions;

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
            Control.SolutionTree.AfterExpand += TreeViewAfterExpandNode;
            Control.SolutionTree.AfterCollapse += TreeViewAfterCollapseNode;
        }

        private void NavigateExplorerTreeNode(object sender, SyntaxTreeNodeClickEventArgs e)
        {
            //todo: re-implement navigate to feature

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
                AddProjectNode(vbProject);
            }
        }

        private void RefreshExplorerTreeView(object sender, System.EventArgs e)
        {
            RefreshExplorerTreeView();
        }

        private void AddProjectNode(VBProject project)
        {
            var treeView = Control.SolutionTree;
            // todo: [re-]implement

            var projectNode = new TreeNode();
            projectNode.Text = project.Name;

            //projectNode.Tag = node.Instruction;
            projectNode.ImageKey = "ClosedFolder";
            treeView.BackColor = treeView.BackColor;

            var moduleNodes = CreateModuleNodes(project, new Font(treeView.Font, FontStyle.Regular));

            projectNode.Nodes.AddRange(moduleNodes.ToArray());
            treeView.Nodes.Add(projectNode);
        }

        private ConcurrentBag<TreeNode> CreateModuleNodes(VBProject project, Font font)
        {
            var moduleNodes = new ConcurrentBag<TreeNode>();


            foreach (VBComponent component in project.VBComponents)
            {
                var moduleNode = new TreeNode(component.Name);
                moduleNode.NodeFont = font;
                moduleNode.ImageKey = GetComponentImageKey(component.Type);

                var parserNode = _parser.Parse(project.Name, component.Name, component.CodeModule.Lines[1, component.CodeModule.CountOfLines]);

                AddNodes<OptionNode>(moduleNode, parserNode, CreateOptionNode);
                AddNodes<EnumNode>(moduleNode, parserNode, CreateEnumNode);
                AddNodes<ProcedureNode>(moduleNode, parserNode, CreateProcedureNode);

                moduleNodes.Add(moduleNode);
            }
            return moduleNodes;
        }

        private delegate TreeNode CreateTreeNode(INode node);
        private void AddNodes<T>(TreeNode moduleNode, INode parserNode, CreateTreeNode createTreeNodeDelegate)
        {
            foreach (INode node in parserNode.Children.OfType<T>())
            {
                var treeNode = createTreeNodeDelegate(node);
                moduleNode.Nodes.Add(createTreeNodeDelegate(node));
            }
        }

        private TreeNode CreateProcedureNode(INode node)
        {
            var result = new TreeNode(node.LocalScope);
            result.ImageKey = GetProcedureImageKey((ProcedureNode)node);

            return result;
        }

        private TreeNode CreateEnumNode(INode node)
        {
            var enumNode = (EnumNode)node;
            var result = new TreeNode(enumNode.Identifier.Name);
            result.ImageKey = enumNode.Accessibility.ToString() + "Enum";

            return result;
        }

        private TreeNode CreateOptionNode(INode node)
        {
            var optionNode = (OptionNode)node;
            var treeNode = new TreeNode("Option" + optionNode.Option);
            treeNode.ImageKey = "Option";

            return treeNode;
        }

        private string GetProcedureImageKey(ProcedureNode node)
        {
            string procKind = string.Empty; //initialize to empty to shut the compiler up
            switch (node.Kind)
            {
                case ProcedureNode.VBProcedureKind.Sub:
                case ProcedureNode.VBProcedureKind.Function:
                    procKind = "Method";
                    break;
                case ProcedureNode.VBProcedureKind.PropertyGet:
                case ProcedureNode.VBProcedureKind.PropertyLet:
                case ProcedureNode.VBProcedureKind.PropertySet:
                    procKind = "Property";
                    break;
            }

            return node.Accessibility.ToString() + procKind;
        }

        private string GetComponentImageKey(vbext_ComponentType componentType)
        {
            //todo: figure out how to get to Interfaces; ImageKey = "PublicInterface"
            switch (componentType)
            {
                case vbext_ComponentType.vbext_ct_ClassModule:
                case vbext_ComponentType.vbext_ct_Document:
                case vbext_ComponentType.vbext_ct_MSForm:
                    return "ClassModule";
                case vbext_ComponentType.vbext_ct_StdModule:
                    return "StandardModule";
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                default:
                    return string.Empty;
            }
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
