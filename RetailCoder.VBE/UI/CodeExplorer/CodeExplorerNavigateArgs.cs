using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.VBA.Grammar;
using Rubberduck.Extensions;

namespace Rubberduck.UI.CodeExplorer
{
    [System.Runtime.InteropServices.ComVisible(false)]
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
