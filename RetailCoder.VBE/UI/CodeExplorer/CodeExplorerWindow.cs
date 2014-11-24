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
            SolutionTree.ShowLines = false;
            SolutionTree.ImageList = TreeNodeIcons;
            SolutionTree.ShowNodeToolTips = true;
            SolutionTree.LabelEdit = false;
        }

        public event EventHandler<SyntaxTreeNodeClickEventArgs> NavigateTreeNode;
        private void SolutionTreeNodeMouseDoubleClicked(object sender, TreeNodeMouseClickEventArgs e)
        {
            var handler = NavigateTreeNode;
            if (handler == null)
            {
                return;
            }

            var instruction = (Instruction)e.Node.Tag;
            handler(this, new SyntaxTreeNodeClickEventArgs(instruction));
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
        public SyntaxTreeNodeClickEventArgs(Instruction instruction)
        {
            _instruction = instruction;
        }

        private readonly Instruction _instruction;
        public Instruction Instruction { get { return _instruction; } }
    }
}
