using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Antlr4.Runtime;

namespace Rubberduck.Inspections.CodePathAnalysis
{
    public class Walker
    {
        public INode GenerateTree(IParseTree tree, Declaration declaration)
        {
            AssignmentNode lastAssignment = null;
            return GenerateTree(tree, declaration, ref lastAssignment);
        }

        public INode GenerateTree(IParseTree tree, Declaration declaration, ref AssignmentNode lastAssignment)
        {
            INode node = default;
            VBAParser.BooleanExpressionContext branchCondition;
            switch (tree)
            {
                case VBAParser.ForNextStmtContext _:
                case VBAParser.ForEachStmtContext _:
                case VBAParser.WhileWendStmtContext _:
                case VBAParser.DoLoopStmtContext _:
                    node = new LoopNode(tree);
                    break;

                case VBAParser.IfStmtContext _:
                    node = new BranchNode(tree);
                    break;
                case VBAParser.ElseBlockContext _:
                    node = new BranchNode(tree);
                    lastAssignment = null;
                    break;
                case VBAParser.ElseIfBlockContext _:
                    node = new BranchNode(tree);
                    break;
                case VBAParser.SingleLineIfStmtContext _:
                    node = new BranchNode(tree);
                    break;
                case VBAParser.SingleLineElseClauseContext _:
                    node = new BranchNode(tree);
                    lastAssignment = null;
                    break;
                case VBAParser.CaseClauseContext _:
                    node = new BranchNode(tree);
                    break;
                case VBAParser.CaseElseClauseContext _:
                    node = new BranchNode(tree);
                    lastAssignment = null;
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
                    node = lastAssignment = new AssignmentNode(tree)
                    {
                        Reference = reference
                    };
                }
                else
                {
                    node = new ReferenceNode(tree, lastAssignment)
                    {
                        Reference = reference
                    };
                    lastAssignment?.AddUsage(node);
                }
            }

            if (node == null)
            {
                node = new GenericNode(tree);
            }

            var children = new HashSet<INode>();
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

                var rhsNode = GenerateTree(rhs, declaration, ref lastAssignment);
                rhsNode.Parent = node;
                children.Add(rhsNode);

                var lhsNode = GenerateTree(lhs, declaration, ref lastAssignment);
                lhsNode.Parent = node;
                children.Add(lhsNode);
            }
            else
            {
                for (var i = 0; i < tree.ChildCount; i++)
                {
                    var nextChild = GenerateTree(tree.GetChild(i), declaration, ref lastAssignment);
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
