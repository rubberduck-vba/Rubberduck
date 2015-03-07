using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.VBA.Grammar;
using Rubberduck.Extensions;

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
            RefreshButton.Click += RefreshButtonClicked;
            SolutionTree.NodeMouseDoubleClick += SolutionTreeNodeMouseDoubleClicked;
            SolutionTree.MouseDown += SolutionTreeMouseDown;
            SolutionTree.AfterExpand += SolutionTreeAfterExpand;
            SolutionTree.AfterCollapse += SolutionTreeAfterCollapse;
            SolutionTree.BeforeExpand += SolutionTreeBeforeExpand;
            SolutionTree.BeforeCollapse += SolutionTreeBeforeCollapse;
            SolutionTree.ShowLines = false;
            SolutionTree.ImageList = TreeNodeIcons;
            SolutionTree.ShowNodeToolTips = true;
            SolutionTree.LabelEdit = false;
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

        #region Hack to disable double click node expansion
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
        #endregion

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
}
