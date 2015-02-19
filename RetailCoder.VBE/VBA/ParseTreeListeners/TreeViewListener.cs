using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public enum TreeViewDisplayStyle
    {
        MemberNames,
        Signatures
    }

    public class TreeViewListener : IVBBaseListener, IExtensionListener<TreeNode>
    {
        private readonly QualifiedModuleName _name;
        private readonly TreeViewDisplayStyle _displayStyle;
        private readonly TreeNode _tree;
        private bool _isInDeclarationsSection = true;

        public TreeViewListener(QualifiedModuleName name, TreeViewDisplayStyle displayStyle = TreeViewDisplayStyle.MemberNames)
        {
            _name = name;
            _displayStyle = displayStyle;
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

            var nodeText = _displayStyle == TreeViewDisplayStyle.Signatures
                ? context.GetText()
                : context.AmbiguousIdentifier().GetText();

            var node = new TreeNode(nodeText);
            var parent = context.Parent as VBParser.VariableStmtContext;
            var accessibility = parent == null || parent.Visibility() == null 
                ? VBAccessibility.Implicit 
                : parent.Visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicField"
                : "PrivateField";

            node.SelectedImageKey = node.ImageKey;
            node.Tag = context.GetQualifiedSelection(_name);
            _tree.Nodes.Add(node);
        }

        public override void EnterConstSubStmt(VBParser.ConstSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var nodeText = _displayStyle == TreeViewDisplayStyle.Signatures
                ? context.GetText()
                : context.AmbiguousIdentifier().GetText();

            var node = new TreeNode(nodeText);
            var parent = context.Parent as VBParser.ConstStmtContext;
            var accessibility = parent == null || parent.Visibility() == null 
                ? VBAccessibility.Implicit 
                : parent.Visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicConst"
                : "PrivateConst";

            node.SelectedImageKey = node.ImageKey;
            node.Tag = context.GetQualifiedSelection(_name);
            _tree.Nodes.Add(node);
        }

        public override void EnterEnumerationStmt(VBParser.EnumerationStmtContext context)
        {
            var node = new TreeNode(context.AmbiguousIdentifier().GetText());
            var members = context.enumerationStmt_Constant();
            foreach (var member in members)
            {
                var memberNode = node.Nodes.Add(member.GetText());
                memberNode.ImageKey = "EnumItem";
                memberNode.SelectedImageKey = memberNode.ImageKey;
                memberNode.Tag = member.GetQualifiedSelection(_name);
            }

            var accessibility = context.Visibility() == null 
                ? VBAccessibility.Implicit
                : context.Visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicEnum"
                : "PrivateEnum";

            node.SelectedImageKey = node.ImageKey;
            node.Tag = context.GetQualifiedSelection(_name);
            _tree.Nodes.Add(node);
        }

        public override void EnterTypeStmt(VBParser.TypeStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.typeStmt_Element();
            foreach (var member in members)
            {
                var memberNodeText = _displayStyle == TreeViewDisplayStyle.Signatures
                    ? member.GetText()
                    : member.AmbiguousIdentifier().GetText();

                var memberNode = node.Nodes.Add(memberNodeText);
                memberNode.ImageKey = "PublicField";
                memberNode.SelectedImageKey = memberNode.ImageKey;
                memberNode.Tag = member.GetQualifiedSelection(_name);
            }

            var accessibility = context.visibility() == null
                ? VBAccessibility.Implicit
                : context.visibility().GetAccessibility();
            node.ImageKey = (accessibility == VBAccessibility.Public || 
                             accessibility == VBAccessibility.Global)
                ? "PublicType"
                : "PrivateType";

            node.Tag = context.GetQualifiedSelection(_name);
            node.SelectedImageKey = node.ImageKey;

            _tree.Nodes.Add(node);
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

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
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

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
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

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
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

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
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

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
        }

        private TreeNode CreateProcedureNode(dynamic context, string imageKey)
        {
            var node = new TreeNode
            {
                ImageKey = imageKey,
                SelectedImageKey = imageKey,
                Tag = ((ParserRuleContext) context).GetQualifiedSelection(_name),
                Text = _displayStyle == TreeViewDisplayStyle.Signatures
                    ? context.Signature()
                    : context.AmbiguousIdentifier().GetText()
            };
            return node;
        }
    }
}
