using System.Windows.Forms;
using Rubberduck.Parsing;

namespace Rubberduck.UI.CodeExplorer
{
    public class TreeNodeNavigateCodeEventArgs : NavigateCodeEventArgs
    {
        private readonly TreeNode _node;
        public TreeNode Node { get { return _node; } }

        public TreeNodeNavigateCodeEventArgs(TreeNode node, QualifiedSelection selection)
            : base(selection)
        {
            _node = node;
        }
    }
}
