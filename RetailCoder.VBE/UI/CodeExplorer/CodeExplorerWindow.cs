using System;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Properties;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.UI.CodeExplorer
{
    public partial class CodeExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "C5318B59-172F-417C-88E3-B377CDA2D809";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return "Code Explorer"; } }

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
            NavigateContextButton.Click += SolutionTreeClick;

            RunAllTestsContextButton.Click += RunAllTestsContextButton_Click;
            InspectContextButton.Click += InspectContextButton_Click;
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

        private void SolutionTreeClick(object sender, EventArgs e)
        {
            var node = SolutionTree.SelectedNode;
            ShowDesignerButton.Enabled = (node != null && node.ImageKey == "Form");
            ShowDesignerContextButton.Enabled = ShowDesignerButton.Enabled;

            SelectedNodeLabel.Text =
                node == null
                    ? string.Empty
                    : node.Text;
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
                var qualifiedSelection = (QualifiedSelection)e.Node.Tag;
                handler(this, new TreeNodeNavigateCodeEventArgs(e.Node, qualifiedSelection));
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
