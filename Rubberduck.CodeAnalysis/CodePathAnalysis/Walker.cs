using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace Rubberduck.Inspections.CodePathAnalysis
{
    public class Walker
    {
        public INode GenerateTree(IParseTree tree, Declaration declaration)
        {
            AssignmentNode lastAssignment = null;
            return GenerateTree(tree, declaration, ref lastAssignment);
        }

        public INode GenerateTree(IParseTree tree, Declaration declaration, ref AssignmentNode lastAssignment, bool isConditional = false, bool isInsideLoop = false)
        {
            INode node = default;
            switch (tree)
            {
                case VBAParser.ForNextStmtContext _:
                case VBAParser.ForEachStmtContext _:
                case VBAParser.WhileWendStmtContext _:
                case VBAParser.DoLoopStmtContext _:
                    node = new LoopNode(tree);
                    isInsideLoop = true;
                    break;
                case VBAParser.IfStmtContext _:
                case VBAParser.ElseBlockContext _:
                case VBAParser.ElseIfBlockContext _:
                case VBAParser.SingleLineIfStmtContext _:
                case VBAParser.SingleLineElseClauseContext _:
                case VBAParser.CaseClauseContext _:
                case VBAParser.CaseElseClauseContext _:
                    node = new BranchNode(tree);
                    isConditional = true;
                    break;
                case VBAParser.BlockContext _:
                    node = new BlockNode(tree);
                    break;
            }

            if (ReferenceEquals(declaration.Context, tree))
            {
                node = new DeclarationNode(tree)
                {
                    Declaration = declaration
                };
            }

            var reference = declaration.References.SingleOrDefault(w => ReferenceEquals(w.Context, tree));
            if (reference != null)
            {
                if (reference.IsAssignment)
                {
                    node = lastAssignment = new AssignmentNode(tree, isConditional, isInsideLoop)
                    {
                        Reference = reference
                    };
                }
                else
                {
                    node = new ReferenceNode(tree, isConditional)
                    {
                        Reference = reference
                    };
                    lastAssignment?.AddUsage(node);
                    isConditional = false;
                }
            }

            if (node == null)
            {
                node = new GenericNode(tree);
            }

            var children = new List<INode>();
            VBAParser.ExpressionContext rhs = null;
            VBAParser.LExpressionContext lhs = null;
            if (tree is VBAParser.LetStmtContext letStmt)
            {
                rhs = letStmt.expression();
                lhs = letStmt.lExpression();
            }
            else if (tree is VBAParser.SetStmtContext setStmt)
            {
                rhs = setStmt.expression();
                lhs = setStmt.lExpression();
            }

            if (rhs != null)
            {
                // add RHS before LHS to match evaluation order

                var rhsNode = GenerateTree(rhs, declaration, ref lastAssignment, isConditional, isInsideLoop);
                rhsNode.Parent = node;
                children.Add(rhsNode);

                var lhsNode = GenerateTree(lhs, declaration, ref lastAssignment, isConditional, isInsideLoop);
                lhsNode.Parent = node;
                children.Add(lhsNode);
            }
            else
            {
                for (var i = 0; i < tree.ChildCount; i++)
                {
                    var nextChild = GenerateTree(tree.GetChild(i), declaration, ref lastAssignment, isConditional, isInsideLoop);
                    nextChild.SortOrder = i;
                    nextChild.Parent = node;

                    if (nextChild.Children.Any() || nextChild.GetType() != typeof(GenericNode))
                    {
                        children.Add(nextChild);
                    }
                }
            }

            node.Children = children.ToImmutableList();
            return node;
        }
    }
}
