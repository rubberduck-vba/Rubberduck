using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.CodeExplorer
{
    public class TreeNodeNavigateCodeEventArgs : NavigateCodeEventArgs
    {
        private readonly TreeNode _node;
        public TreeNode Node { get { return _node; } }

        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        public TreeNodeNavigateCodeEventArgs(TreeNode node, Declaration declaration)
            : this(node, new QualifiedSelection(declaration.QualifiedName.QualifiedModuleName, declaration.Selection))
        {
            _declaration = declaration;
        }

        public TreeNodeNavigateCodeEventArgs(TreeNode node, QualifiedSelection selection)
            : base(selection)
        {
            _node = node;
        }
    }
}
