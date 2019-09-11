using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using VBAParser = Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.CodePathAnalysis
{
    public static class NodeExtensions
    {
        public static void SetCurrentNode(this IParseTree tree, ProcedureTreeVisitorState state)
        {
            switch (tree)
            {
                case VBAParser.ForNextStmtContext _:
                case VBAParser.ForEachStmtContext _:
                case VBAParser.WhileWendStmtContext _:
                case VBAParser.DoLoopStmtContext _:
                    state.OnLoopNode(tree);
                    break;

                case VBAParser.IfStmtContext _:
                case VBAParser.ElseBlockContext _:
                case VBAParser.ElseIfBlockContext _:
                case VBAParser.SingleLineIfStmtContext _:
                case VBAParser.SingleLineElseClauseContext _:
                case VBAParser.CaseClauseContext _:
                case VBAParser.CaseElseClauseContext _:
                    state.OnBranchNode(tree);
                    break;

                case VBAParser.BlockContext _:
                    state.OnBlockNode(tree);
                    break;

                case VBAParser.BlockStmtContext _:
                    state.OnStatementNode(tree);
                    break;

                case VBAParser.GoToStmtContext stmt:
                    state.OnGoToNode(stmt);
                    break;
                case VBAParser.GoSubStmtContext stmt:
                    state.OnGoSubNode(stmt);
                    break;
                case VBAParser.ReturnStmtContext stmt:
                    state.OnReturnNode(stmt);
                    break;

                case VBAParser.OnErrorStmtContext stmt:
                    state.OnErrorNode(stmt);
                    break;
                case VBAParser.ResumeStmtContext stmt:
                    state.OnResume(stmt);
                    break;

                default:
                    state.CurrentNode = new GenericNode(tree);
                    break;
            }
        }
    }

    public class ProcedureTreeVisitor
    {
        public INode GenerateTree(QualifiedMemberName scope, ProcedureTreeVisitorState state, IParseTree tree = null)
        {
            if (tree == null)
            {
                // debug: we want the top-level parser rule for the procedure (SubStmt, FunctionStmt, etc.)
                tree = state.Declarations.Single(d => d.QualifiedName.Equals(scope)).Context.Parent;
            }
            
            tree.SetCurrentNode(state);
            var declaration = state.Declarations.SingleOrDefault(d => ReferenceEquals(d.Context, tree));
            if (declaration != null)
            {
                state.CurrentNode = new DeclarationNode(tree) { Declaration = declaration };
            }
            else
            {
                var reference = state.IdentifierReferences.SingleOrDefault(r => ReferenceEquals(r.Context, tree));
                if (reference != null)
                {
                    if (reference.IsAssignment)
                    {
                        var node = new AssignmentNode(tree) { Reference = reference };
                        state.OnAssignmentNode(node);
                    }
                    else
                    {
                        var node = new ReferenceNode(tree) { Reference = reference, };
                        state.OnReferenceNode(node);
                    }
                }
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

                var rhsNode = GenerateTree(scope, state, rhs);
                rhsNode.Parent = state.CurrentNode;
                children.Add(rhsNode);

                var lhsNode = GenerateTree(scope, state, lhs);
                lhsNode.Parent = state.CurrentNode;
                children.Add(lhsNode);
            }
            else
            {
                for (var i = 0; i < tree.ChildCount; i++)
                {
                    var nextChild = GenerateTree(scope, state, tree.GetChild(i));
                    nextChild.SortOrder = i;
                    nextChild.Parent = state.CurrentNode;

                    if (nextChild.Children.Any() || nextChild.GetType() != typeof(GenericNode))
                    {
                        children.Add(nextChild);
                    }
                }
            }

            state.CurrentNode.Children = children.ToList();
            return state.CurrentNode;
        }        
    }
}
