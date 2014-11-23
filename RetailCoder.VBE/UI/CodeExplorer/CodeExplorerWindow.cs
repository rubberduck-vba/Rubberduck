using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeExplorer
{
    [ComVisible(false)]
    public partial class CodeExplorerWindow : UserControl
    {
        public CodeExplorerWindow()
        {
            InitializeComponent();
            RefreshButton.Click += RefreshButtonClicked;
            SolutionTree.NodeMouseDoubleClick += SolutionTreeNodeMouseDoubleClicked;
        }

        public event EventHandler<SyntaxTreeNodeClickEventArgs> NavigateTreeNode;
        private void SolutionTreeNodeMouseDoubleClicked(object sender, TreeNodeMouseClickEventArgs e)
        {
            var handler = NavigateTreeNode;
            if (handler == null)
            {
                return;
            }

            var node = e.Node.Tag as SyntaxTreeNode;
            handler(this, new SyntaxTreeNodeClickEventArgs(node));
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

    public class SyntaxTreeNodeClickEventArgs : EventArgs
    {
        public SyntaxTreeNodeClickEventArgs(SyntaxTreeNode node)
        {
            _node = node;
        }

        private readonly SyntaxTreeNode _node;
        public SyntaxTreeNode Node { get { return _node; } }
    }
}
