using System.Windows.Forms;
using Rubberduck.Extensions;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerNavigateArgs : NavigateCodeEventArgs
    {
        private readonly TreeNode node;
        public TreeNode Node { get { return node; } }

        public CodeExplorerNavigateArgs(TreeNode node, QualifiedSelection selection)
            : base(selection)
        {
            this.node = node;
        }
    }
}
