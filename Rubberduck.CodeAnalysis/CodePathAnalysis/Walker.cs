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
            INode node = default;
            switch (tree)
            {
                case VBAParser.ForNextStmtContext _:
                case VBAParser.ForEachStmtContext _:
                case VBAParser.WhileWendStmtContext _:
                case VBAParser.DoLoopStmtContext _:
                    node = new LoopNode(tree);
                    break;
                case VBAParser.IfStmtContext _:
                case VBAParser.ElseBlockContext _:
                case VBAParser.ElseIfBlockContext _:
                case VBAParser.SingleLineIfStmtContext _:
                case VBAParser.SingleLineElseClauseContext _:
                case VBAParser.CaseClauseContext _:
                case VBAParser.CaseElseClauseContext _:
                    node = new BranchNode(tree);
                    break;
                case VBAParser.BlockContext _:
                    node = new BlockNode(tree);
                    break;
                case VBAParser.LetStmtContext _:
                case VBAParser.SetStmtContext _:
                    node = new AssignmentExpressionNode(tree);
                    break;
            }

            if (declaration.Context == tree)
            {
                node = new DeclarationNode(tree)
                {
                    Declaration = declaration
                };
            }

            var reference = declaration.References.SingleOrDefault(w => w.Context == tree);
            if (reference != null)
            {
                if (reference.IsAssignment)
                {
                    node = new AssignmentNode(tree)
                    {
                        Reference = reference
                    };
                }
                else
                {
                    node = new ReferenceNode(tree)
                    {
                        Reference = reference
                    };
                }
            }

            if (node == null)
            {
                node = new GenericNode(tree);
            }

            var children = new List<INode>();
            for (var i = 0; i < tree.ChildCount; i++)
            {
                var nextChild = GenerateTree(tree.GetChild(i), declaration);
                nextChild.SortOrder = i;
                nextChild.Parent = node;

                if (nextChild.Children.Any() || nextChild.GetType() != typeof(GenericNode))
                {
                    children.Add(nextChild);
                }
            }

            node.Children = children.ToImmutableList();

            return node;
        }
    }
}
