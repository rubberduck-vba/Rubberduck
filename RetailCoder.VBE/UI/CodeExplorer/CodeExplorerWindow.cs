using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.VBA.Grammar;
using Rubberduck.Extensions;

namespace Rubberduck.UI.CodeExplorer
{
    [ComVisible(false)]
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
            SolutionTree.BeforeExpand += SolutionTreeBeforeExpand;
            SolutionTree.BeforeCollapse += SolutionTreeBeforeCollapse;
            SolutionTree.ShowLines = false;
            SolutionTree.ImageList = TreeNodeIcons;
            SolutionTree.ShowNodeToolTips = true;
            SolutionTree.LabelEdit = false;
        }

        #region Hack to disable double click node expansion
        private bool doubleClicked;
        private void SolutionTreeMouseDown(object sender, MouseEventArgs e)
        {
            doubleClicked = (e.Clicks > 1);
        }

        private void SolutionTreeBeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = doubleClicked;
            doubleClicked = false;
        }

        private void SolutionTreeBeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = doubleClicked;
            doubleClicked = false;
        }
        #endregion

        public event EventHandler<CodeExplorerNavigateArgs> NavigateTreeNode;
        private void SolutionTreeNodeMouseDoubleClicked(object sender, TreeNodeMouseClickEventArgs e)
        {
            var handler = NavigateTreeNode;
            if (handler == null)
            {
                return;
            }

            if (e.Node.Tag != null)
            {
                var qualifiedSelection = (QualifiedSelection)e.Node.Tag;
                handler(this, new CodeExplorerNavigateArgs(e.Node, qualifiedSelection));
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
