using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.CodeExplorer
{
    public class TreeNodeNavigateCodeEventArgs : NavigateCodeEventArgs
    {
        private readonly TreeNode _node;
        public TreeNode Node { get { return _node; } }

        public TreeNodeNavigateCodeEventArgs(TreeNode node)
            : base(node.Tag as Declaration)
        {
            _node = node;
        }

        public TreeNodeNavigateCodeEventArgs(TreeNode node, QualifiedSelection selection)
            : base(selection)
        {
            _node = node;
        }
    }
}
