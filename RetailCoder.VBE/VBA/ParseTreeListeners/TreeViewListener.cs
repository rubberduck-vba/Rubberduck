using System.Collections.Generic;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public enum TreeViewDisplayStyle
    {
        MemberNames,
        Signatures
    }

    public class TreeViewListener : VBABaseListener, IExtensionListener<TreeNode>
    {
        private readonly QualifiedModuleName _name;
        private readonly TreeViewDisplayStyle _displayStyle;
        private readonly TreeNode _tree;
        private bool _isInDeclarationsSection = true;

        public TreeViewListener(QualifiedModuleName name, TreeViewDisplayStyle displayStyle = TreeViewDisplayStyle.MemberNames)
        {
            _name = name;
            _displayStyle = displayStyle;
            _tree = new TreeNode(name.ModuleName) {Tag = new QualifiedSelection(_name, Selection.Home)};
        }

        public IEnumerable<QualifiedContext<TreeNode>> Members
        {
            get { return new[] {new QualifiedContext<TreeNode>(_name , _tree)}; }
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var nodeText = _displayStyle == TreeViewDisplayStyle.Signatures
                ? context.GetText()
                : context.ambiguousIdentifier().GetText();

            var node = new TreeNode(nodeText);
            var parent = context.Parent as VBAParser.VariableStmtContext;
            var accessibility = parent == null || parent.visibility() == null 
                ? Accessibility.Implicit 
                : parent.visibility().GetAccessibility();
            node.ImageKey = (accessibility == Accessibility.Public || 
                             accessibility == Accessibility.Global)
                ? "PublicField"
                : "PrivateField";

            node.SelectedImageKey = node.ImageKey;
            node.Tag = context.GetQualifiedSelection(_name);
            _tree.Nodes.Add(node);
        }

        public override void EnterConstSubStmt(VBAParser.ConstSubStmtContext context)
        {
            if (!_isInDeclarationsSection)
            {
                return;
            }

            var nodeText = _displayStyle == TreeViewDisplayStyle.Signatures
                ? context.GetText()
                : context.ambiguousIdentifier().GetText();

            var node = new TreeNode(nodeText);
            var parent = context.Parent as VBAParser.ConstStmtContext;
            var accessibility = parent == null || parent.visibility() == null 
                ? Accessibility.Implicit 
                : parent.visibility().GetAccessibility();
            node.ImageKey = (accessibility == Accessibility.Public || 
                             accessibility == Accessibility.Global)
                ? "PublicConst"
                : "PrivateConst";

            node.SelectedImageKey = node.ImageKey;
            node.Tag = context.GetQualifiedSelection(_name);
            _tree.Nodes.Add(node);
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.enumerationStmt_Constant();
            foreach (var member in members)
            {
                var memberNode = node.Nodes.Add(member.GetText());
                memberNode.ImageKey = "EnumItem";
                memberNode.SelectedImageKey = memberNode.ImageKey;
                memberNode.Tag = member.GetQualifiedSelection(_name);
            }

            var accessibility = context.visibility() == null 
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            node.ImageKey = (accessibility == Accessibility.Public || 
                             accessibility == Accessibility.Global)
                ? "PublicEnum"
                : "PrivateEnum";

            node.SelectedImageKey = node.ImageKey;
            node.Tag = context.GetQualifiedSelection(_name);
            _tree.Nodes.Add(node);
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            var node = new TreeNode(context.ambiguousIdentifier().GetText());
            var members = context.typeStmt_Element();
            foreach (var member in members)
            {
                var memberNodeText = _displayStyle == TreeViewDisplayStyle.Signatures
                    ? member.GetText()
                    : member.ambiguousIdentifier().GetText();

                var memberNode = node.Nodes.Add(memberNodeText);
                memberNode.ImageKey = "PublicField";
                memberNode.SelectedImageKey = memberNode.ImageKey;
                memberNode.Tag = member.GetQualifiedSelection(_name);
            }

            var accessibility = context.visibility() == null
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            node.ImageKey = (accessibility == Accessibility.Public || 
                             accessibility == Accessibility.Global)
                ? "PublicType"
                : "PrivateType";

            node.Tag = context.GetQualifiedSelection(_name);
            node.SelectedImageKey = node.ImageKey;

            _tree.Nodes.Add(node);
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.visibility() == null
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            var imageKey = accessibility == Accessibility.Private
                ? "PrivateMethod"
                : accessibility == Accessibility.Friend
                    ? "FriendMethod"
                    : "PublicMethod";

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.visibility() == null
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            var imageKey = accessibility == Accessibility.Private
                ? "PrivateMethod"
                : accessibility == Accessibility.Friend
                    ? "FriendMethod"
                    : "PublicMethod";

            var node = CreateProcedureNode(context, imageKey);
            _tree.Nodes.Add(node);
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.visibility() == null
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            var imageKey = accessibility == Accessibility.Private
                ? "PrivateProperty"
                : accessibility == Accessibility.Friend
                    ? "FriendProperty"
                    : "PublicProperty";

            var node = CreateProcedureNode(context, imageKey);
            if (_displayStyle == TreeViewDisplayStyle.MemberNames)
            {
                node.Text += (" (" + Tokens.Get + ")");
            }
            _tree.Nodes.Add(node);
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.visibility() == null
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            var imageKey = accessibility == Accessibility.Private
                ? "PrivateProperty"
                : accessibility == Accessibility.Friend
                    ? "FriendProperty"
                    : "PublicProperty";

            var node = CreateProcedureNode(context, imageKey);
            if (_displayStyle == TreeViewDisplayStyle.MemberNames)
            {
                node.Text += (" (" + Tokens.Let + ")");
            }
            _tree.Nodes.Add(node);
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _isInDeclarationsSection = false;
            var accessibility = context.visibility() == null
                ? Accessibility.Implicit
                : context.visibility().GetAccessibility();
            var imageKey = accessibility == Accessibility.Private
                ? "PrivateProperty"
                : accessibility == Accessibility.Friend
                    ? "FriendProperty"
                    : "PublicProperty";

            var node = CreateProcedureNode(context, imageKey);
            if (_displayStyle == TreeViewDisplayStyle.MemberNames)
            {
                node.Text += (" (" + Tokens.Set + ")");
            }
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
                    ? ParserRuleContextExtensions.Signature(context)
                    : context.ambiguousIdentifier().GetText()
            };
            return node;
        }
    }
}
