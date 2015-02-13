using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class TreeViewListener : VisualBasic6BaseListener, IExtensionListener<TreeNode>
    {
        private readonly TreeNode _tree;
        private bool _isInDeclarationsSection = true;

        public TreeViewListener(QualifiedModuleName name)
        {
            _tree = new TreeNode(name.ModuleName);
        }

        public IEnumerable<TreeNode> Members
        {
            get { return new[] {_tree}; }
        }

        public override void EnterVariableSubStmt(VisualBasic6Parser.VariableSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var node = new TreeNode(context.GetText());
            _tree.Nodes.Add(node);
        }

        public override void EnterConstSubStmt(VisualBasic6Parser.ConstSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var node = new TreeNode(context.GetText());
            _tree.Nodes.Add(node);
        }

        public override void EnterEnumerationStmt(VisualBasic6Parser.EnumerationStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.enumerationStmt_Constant();
            foreach (var member in members)
            {
                var memberNode = node.Nodes.Add(member.GetText());
                // format node
            }

            _tree.Nodes.Add(node);
        }

        public override void EnterTypeStmt(VisualBasic6Parser.TypeStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.typeStmt_Element();
            foreach (var member in members)
            {
                var memberNode = node.Nodes.Add(member.GetText());
                // format node
            }
        }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            _isInDeclarationsSection = false;
            _tree.Nodes.Add(CreateProcedureNode(context));
        }

        public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            _isInDeclarationsSection = false;
            _tree.Nodes.Add(CreateProcedureNode(context));
        }

        public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            _isInDeclarationsSection = false;
            _tree.Nodes.Add(CreateProcedureNode(context));
        }

        public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            _isInDeclarationsSection = false;
            _tree.Nodes.Add(CreateProcedureNode(context));
        }

        public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            _isInDeclarationsSection = false;
            _tree.Nodes.Add(CreateProcedureNode(context));
        }

        private TreeNode CreateProcedureNode(dynamic context)
        {
            var procedureName = context.ambiguousIdentifier().GetText();
            var node = new TreeNode(procedureName);

            var args = context.argList().arg() as IReadOnlyList<VisualBasic6Parser.ArgContext>;
            if (args == null)
            {
                return node;
            }

            foreach (var arg in args)
            {
                var argNode = new TreeNode(arg.GetText());
                node.Nodes.Add(argNode);
            }

            return node;
        }
    }
}
