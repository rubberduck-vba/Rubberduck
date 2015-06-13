﻿using System;
using System.Windows.Forms;
using NetOffice.VBIDEApi.Enums;

using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;

namespace Rubberduck.UI.CodeExplorer
{
    public partial class CodeExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B59-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeExplorerDockablePresenter_Caption; } }

        public CodeExplorerWindow()
        {
            InitializeComponent();

            ShowFoldersToggleButton.Click += ShowFoldersToggleButtonClick;
            ShowDesignerButton.Click += ShowDesignerButtonClick;
            ShowDesignerContextButton.Click += ShowDesignerButtonClick;
            AddClassButton.Click += AddClassButton_Click;
            AddStdModuleButton.Click += AddStdModuleButton_Click;
            AddFormButton.Click += AddFormButton_Click;
            AddTestModuleButton.Click += AddTestModuleButtonClick;

            DisplayMemberNamesButton.Click += DisplayMemberNamesButton_Click;
            DisplaySignaturesButton.Click += DisplaySignaturesButton_Click;

            RefreshButton.Click += RefreshButtonClicked;
            RefreshContextButton.Click += RefreshButtonClicked;
            SolutionTree.NodeMouseDoubleClick += SolutionTreeNodeMouseDoubleClicked;
            SolutionTree.AfterExpand += SolutionTreeAfterExpand;
            SolutionTree.AfterCollapse += SolutionTreeAfterCollapse;
            SolutionTree.AfterSelect += SolutionTreeClick;
            SolutionTree.MouseDown += SolutionTreeMouseDown;
            SolutionTree.BeforeExpand += SolutionTreeBeforeExpand;
            SolutionTree.BeforeCollapse += SolutionTreeBeforeCollapse;
            SolutionTree.ShowLines = false;
            SolutionTree.ImageList = TreeNodeIcons;
            SolutionTree.ShowNodeToolTips = true;
            SolutionTree.LabelEdit = false;

            AddClassContextButton.Click += AddClassButton_Click;
            AddStdModuleContextButton.Click += AddStdModuleButton_Click;
            AddFormContextButton.Click += AddFormButton_Click;
            AddTestModuleContextButton.Click += AddTestModuleButtonClick;
            NavigateContextButton.Click += NavigateContextButtonClick;
            RenameContextButton.Click += RenameContextButtonClick;

            RunAllTestsContextButton.Click += RunAllTestsContextButton_Click;
            InspectContextButton.Click += InspectContextButton_Click;
            FindAllReferencesContextButton.Click += FindAllReferencesContextButton_Click;
            FindAllImplementationsContextButton.Click += FindAllImplementationsContextButton_Click;
            
            RefreshButton.ToolTipText = RubberduckUI.Refresh;
            ShowFoldersToggleButton.ToolTipText = RubberduckUI.CodeExplorer_ShowFoldersToolTip;
            ShowDesignerButton.ToolTipText = RubberduckUI.CodeExplorer_ShowDesignerToolTip;

            AddClassButton.Text = RubberduckUI.CodeExplorer_AddClassText;
            AddStdModuleButton.Text = RubberduckUI.CodeExplorer_AddStdModuleText;
            AddFormButton.Text = RubberduckUI.CodeExplorer_AddFormText;
            AddTestModuleButton.Text = RubberduckUI.CodeExplorer_AddTestModuleText;
            DisplayMemberNamesButton.Text = RubberduckUI.CodeExplorer_DisplayMemberNamesText;
            DisplaySignaturesButton.Text = RubberduckUI.CodeExplorer_DisplaySignaturesText;

            AddButton.Text = RubberduckUI.CodeExplorer_New;
            newToolStripMenuItem.Text = RubberduckUI.New;

            RefreshContextButton.Text = RubberduckUI.CodeExplorer_Refresh;
            ShowDesignerContextButton.Text = RubberduckUI.CodeExplorer_ShowDesignerText;
            AddClassContextButton.Text = RubberduckUI.CodeExplorer_AddClassText;
            AddStdModuleContextButton.Text = RubberduckUI.CodeExplorer_AddStdModuleText;
            AddFormContextButton.Text = RubberduckUI.CodeExplorer_AddFormText;
            AddTestModuleContextButton.Text = RubberduckUI.CodeExplorer_AddTestModuleText;
            NavigateContextButton.Text = RubberduckUI.Navigate;
            RenameContextButton.Text = RubberduckUI.Rename;
            RunAllTestsContextButton.Text = RubberduckUI.CodeExplorer_RunAllTestsText;
            InspectContextButton.Text = RubberduckUI.Inspect;
            FindAllReferencesContextButton.Text = RubberduckUI.CodeExplorer_FindAllReferencesText;
            FindAllImplementationsContextButton.Text = RubberduckUI.CodeExplorer_FindAllImplementationsText;
        }

        public void EnableRefresh(bool enabled = true)
        {
            RefreshButton.Enabled = enabled;
        }

        public event EventHandler<NavigateCodeEventArgs> FindAllReferences;
        private void FindAllReferencesContextButton_Click(object sender, EventArgs e)
        {
            var handler = FindAllReferences;
            if (handler != null && SolutionTree.SelectedNode != null)
            {
                var target = SolutionTree.SelectedNode.Tag as Declaration;
                if (target != null)
                {
                    handler(this, new NavigateCodeEventArgs(target));
                }
            }
        }

        public event EventHandler<NavigateCodeEventArgs> FindAllImplementations;
        private void FindAllImplementationsContextButton_Click(object sender, EventArgs e)
        {
            var handler = FindAllImplementations;
            if (handler != null && SolutionTree.SelectedNode != null)
            {
                var target = SolutionTree.SelectedNode.Tag as Declaration;
                if (target != null)
                {
                    handler(this, new NavigateCodeEventArgs(target));
                }
            }
        }

        public event EventHandler<TreeNodeNavigateCodeEventArgs> Rename;
        private void RenameContextButtonClick(object sender, EventArgs e)
        {
            var handler = Rename;
            if (handler != null && SolutionTree.SelectedNode != null)
            {
                handler(this, new TreeNodeNavigateCodeEventArgs(SolutionTree.SelectedNode));
            }
        }

        private void NavigateContextButtonClick(object sender, EventArgs e)
        {
            var handler = NavigateTreeNode;
            if (handler != null && SolutionTree.SelectedNode != null)
            {
                handler(this, new TreeNodeNavigateCodeEventArgs(SolutionTree.SelectedNode));
            }
        }

        public event EventHandler RunInspections;
        private void InspectContextButton_Click(object sender, EventArgs e)
        {
            var handler = RunInspections;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler RunAllTests;
        private void RunAllTestsContextButton_Click(object sender, EventArgs e)
        {
            var handler = RunAllTests;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private TreeViewDisplayStyle _displayStyle;
        public event EventHandler DisplayStyleChanged;

        public TreeViewDisplayStyle DisplayStyle
        {
            get { return _displayStyle; }
            private set
            {
                _displayStyle = value;

                var handler = DisplayStyleChanged;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
            }
        }

        private void DisplaySignaturesButton_Click(object sender, EventArgs e)
        {
            DisplayStyle = TreeViewDisplayStyle.Signatures;
            CheckDisplayStyleButton();
        }

        private void DisplayMemberNamesButton_Click(object sender, EventArgs e)
        {
            DisplayStyle = TreeViewDisplayStyle.MemberNames;
            CheckDisplayStyleButton();
        }

        private void CheckDisplayStyleButton()
        {
            DisplaySignaturesButton.Checked = DisplayStyle == TreeViewDisplayStyle.Signatures;
            DisplayMemberNamesButton.Checked = DisplayStyle == TreeViewDisplayStyle.MemberNames;
            DisplayModeButton.Image = DisplayStyle == TreeViewDisplayStyle.Signatures
                ? Resources.DisplayFullSignature_13393_32
                : Resources.DisplayName_13394_32;
        }

        public event EventHandler AddTestModule;
        private void AddTestModuleButtonClick(object sender, EventArgs e)
        {
            var handler = AddTestModule;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler<TreeNodeNavigateCodeEventArgs> SelectionChanged;
        private void SolutionTreeClick(object sender, EventArgs e)
        {
            var node = SolutionTree.SelectedNode;
            ShowDesignerButton.Enabled = (node != null && node.ImageKey == "Form");
            ShowDesignerContextButton.Enabled = ShowDesignerButton.Enabled;

            SelectedNodeLabel.Text =
                node == null
                    ? string.Empty
                    : node.Text;

            var handler = SelectionChanged;
            if (handler == null)
            {
                return;
            }

            handler(this, new TreeNodeNavigateCodeEventArgs(node));
        }

        private bool CanDeleteNode(TreeNode node)
        {
            return (node != null && !node.ImageKey.Contains("Folder"));
        }

        public event EventHandler ShowDesigner;
        private void ShowDesignerButtonClick(object sender, EventArgs e)
        {
            var handler = ShowDesigner;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler ToggleFolders;
        private void ShowFoldersToggleButtonClick(object sender, EventArgs e)
        {
            var handler = ToggleFolders;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
            ShowFoldersToggleButton.Checked = !ShowFoldersToggleButton.Checked;
        }

        public event EventHandler<AddComponentEventArgs> AddComponent;

        private void AddFormButton_Click(object sender, EventArgs e)
        {
            var handler = AddComponent;
            if (handler != null)
            {
                handler(this, new AddComponentEventArgs(vbext_ComponentType.vbext_ct_MSForm));
            }
        }

        private void AddStdModuleButton_Click(object sender, EventArgs e)
        {
            var handler = AddComponent;
            if (handler != null)
            {
                handler(this, new AddComponentEventArgs(vbext_ComponentType.vbext_ct_StdModule));
            }
        }

        private void AddClassButton_Click(object sender, EventArgs e)
        {
            var handler = AddComponent;
            if (handler != null)
            {
                handler(this, new AddComponentEventArgs(vbext_ComponentType.vbext_ct_ClassModule));
            }
        }

        private void SolutionTreeAfterCollapse(object sender, TreeViewEventArgs e)
        {
            if (!e.Node.ImageKey.Contains("Folder"))
            {
                return;
            }

            e.Node.ImageKey = "ClosedFolder";
            e.Node.SelectedImageKey = e.Node.ImageKey;
        }

        private void SolutionTreeAfterExpand(object sender, TreeViewEventArgs e)
        {
            if (!e.Node.ImageKey.Contains("Folder"))
            {
                return;
            }

            e.Node.ImageKey = "OpenFolder";
            e.Node.SelectedImageKey = e.Node.ImageKey;
        }

        private bool _doubleClicked;
        private void SolutionTreeMouseDown(object sender, MouseEventArgs e)
        {
            _doubleClicked = (e.Clicks > 1);
        }

        private void SolutionTreeBeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = _doubleClicked;
            if (_doubleClicked && NavigateTreeNode != null)
            {
                //NavigateTreeNode(sender, new TreeNodeNavigateCodeEventArgs(e.Node, (QualifiedSelection)e.Node.Tag));
            }
            _doubleClicked = false;
        }

        private void SolutionTreeBeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = _doubleClicked;
            if (_doubleClicked && NavigateTreeNode != null)
            {
                //NavigateTreeNode(sender, new TreeNodeNavigateCodeEventArgs(e.Node, (QualifiedSelection)e.Node.Tag));
            }
            _doubleClicked = false;
        }

        public event EventHandler<TreeNodeNavigateCodeEventArgs> NavigateTreeNode;
        private void SolutionTreeNodeMouseDoubleClicked(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.ImageKey.Contains("Folder"))
            {
                e.Node.Toggle();
            }

            var handler = NavigateTreeNode;
            if (handler == null)
            {
                return;
            }

            if (e.Node.Tag != null)
            {
                handler(this, new TreeNodeNavigateCodeEventArgs(e.Node));
            }
        }


        public event EventHandler RefreshTreeView;
        private void RefreshButtonClicked(object sender, EventArgs e)
        {
            var handler = RefreshTreeView;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }
    }

    public class AddComponentEventArgs : EventArgs
    {
        public AddComponentEventArgs(vbext_ComponentType type)
        {
            _type = type;
        }

        private readonly vbext_ComponentType _type;
        public vbext_ComponentType ComponentType { get { return _type; } }
    }
}
