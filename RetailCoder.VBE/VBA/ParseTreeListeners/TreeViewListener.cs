using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class TreeViewListener : IVBBaseListener, IExtensionListener<TreeNode>
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

        public override void EnterVariableSubStmt(VBParser.VariableSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var node = new TreeNode(context.GetText());
            var parent = context.Parent as VBParser.VariableStmtContext;
            var accessibility = parent == null || parent.Visibility() == null 
                ? VBAccessibility.Implicit 
                : parent.Visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicField"
                : "PrivateField";

            node.SelectedImageKey = node.ImageKey;
            _tree.Nodes.Add(node);
        }

        public override void EnterConstSubStmt(VBParser.ConstSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var node = new TreeNode(context.GetText());
            var parent = context.Parent as VBParser.ConstStmtContext;
            var accessibility = parent == null || parent.Visibility() == null 
                ? VBAccessibility.Implicit 
                : parent.Visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicConst"
                : "PrivateConst";

            node.SelectedImageKey = node.ImageKey;
            _tree.Nodes.Add(node);
        }

        public override void EnterEnumerationStmt(VBParser.EnumerationStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.enumerationStmt_Constant();
            foreach (var member in members)
            {
                var memberNode = node.Nodes.Add(member.GetText());
                memberNode.ImageKey = "EnumItem";
                memberNode.SelectedImageKey = memberNode.ImageKey;
            }

            var accessibility = context.visibility() == null 
                ? VBAccessibility.Implicit
                : context.visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicEnum"
                : "PrivateEnum";

            node.SelectedImageKey = node.ImageKey;

            _tree.Nodes.Add(node);
        }

        public override void EnterTypeStmt(VBParser.TypeStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.typeStmt_Element();
            foreach (var member in members)
            {
                var memberNode = node.Nodes.Add(member.GetText());
                memberNode.ImageKey = "PublicField";
                memberNode.SelectedImageKey = memberNode.ImageKey;
            }

            var accessibility = context.visibility() == null
                ? VBAccessibility.Implicit
                : context.visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicType"
                : "PrivateType";

            node.SelectedImageKey = node.ImageKey;
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.Visibility() == null
                ? VBAccessibility.Implicit
                : context.Visibility().GetAccessibility();
            var imageKey = accessibility == VBAccessibility.Private
                ? "PrivateMethod"
                : accessibility == VBAccessibility.Friend
                    ? "FriendMethod"
                    : "PublicMethod";
            _tree.Nodes.Add(CreateProcedureNode(context, imageKey));
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.Visibility() == null
                ? VBAccessibility.Implicit
                : context.Visibility().GetAccessibility();
            var imageKey = accessibility == VBAccessibility.Private
                ? "PrivateMethod"
                : accessibility == VBAccessibility.Friend
                    ? "FriendMethod"
                    : "PublicMethod";
            _tree.Nodes.Add(CreateProcedureNode(context, imageKey));
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.visibility() == null
                ? VBAccessibility.Implicit
                : context.visibility().GetAccessibility();
            var imageKey = accessibility == VBAccessibility.Private
                ? "PrivateProperty"
                : accessibility == VBAccessibility.Friend
                    ? "FriendProperty"
                    : "PublicProperty";
            _tree.Nodes.Add(CreateProcedureNode(context, imageKey));
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.Visibility() == null
                ? VBAccessibility.Implicit
                : context.Visibility().GetAccessibility();
            var imageKey = accessibility == VBAccessibility.Private
                ? "PrivateProperty"
                : accessibility == VBAccessibility.Friend
                    ? "FriendProperty"
                    : "PublicProperty";
            _tree.Nodes.Add(CreateProcedureNode(context, imageKey));
        }

        public override void EnterPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.Visibility() == null
                ? VBAccessibility.Implicit
                : context.Visibility().GetAccessibility();
            var imageKey = accessibility == VBAccessibility.Private
                ? "PrivateProperty"
                : accessibility == VBAccessibility.Friend
                    ? "FriendProperty"
                    : "PublicProperty";
            _tree.Nodes.Add(CreateProcedureNode(context,imageKey));
        }

        private TreeNode CreateProcedureNode(dynamic context, string imageKey)
        {
            var procedureName = context.ambiguousIdentifier().GetText();
            var node = new TreeNode(procedureName);

            var args = context.argList().arg() as IReadOnlyList<VBParser.ArgContext>;
            if (args == null)
            {
                return node;
            }

            foreach (var arg in args)
            {
                var argNode = new TreeNode(arg.GetText());
                argNode.ImageKey = "Parameter";
                argNode.SelectedImageKey = argNode.ImageKey;

                node.Nodes.Add(argNode);
            }

            node.ImageKey = imageKey;
            node.SelectedImageKey = node.ImageKey;
            return node;
        }
    }
}
