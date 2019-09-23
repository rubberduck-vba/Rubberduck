using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;
using Rubberduck.Parsing.Grammar;
using System;
using System.Diagnostics;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution.ExtendedNodeVisitor
{
    public class ExtendedNodeVisitor
    {
        private readonly IExtendedNode[] _nodes;
        private readonly HashSet<ILabelNode> _labels;

        public ExtendedNodeVisitor(ModuleBodyElementDeclaration member)
        {
            _nodes = member.Context.FlattenExtendedNodes().ToArray();
            _labels = new HashSet<ILabelNode>(_nodes.OfType<ILabelNode>());
        }

        private readonly List<IExtendedNode> _traversed = new List<IExtendedNode>();

        private readonly List<CodePath> _allPaths = new List<CodePath>();
        private readonly Stack<CodePath> _currentPath = new Stack<CodePath>();

        private int _position = 0;

        public CodePath[] GetAllCodePaths()
        {
            _position = 0;
            EnterExecutionPath();

            while (_position < _nodes.Length)
            {
                var node = _nodes[_position];
                if (node == _traversed.LastOrDefault())
                {
                    // which is more likely:
                    // a self-referencing goto-loop, or a bug somewhere?
                    Debug.Assert(false);
                }
                _traversed.Add(node);
                _allPaths.AddRange(VisitExtendedNode(node));
            }

            return _allPaths.ToArray();
        }

        private CodePath[] VisitExtendedNode(IExtendedNode node)
        {
            HitNode(node);
            switch(node)
            {
                case IBranchNode branchNode:
                    return VisitExtendedNode(branchNode);
                case IJumpNode jumpNode:
                    return VisitExtendedNode(jumpNode);
                case IExecutableNode exeNode:
                    return VisitExtendedNode(exeNode);
                case IEvaluatableNode evalNode:
                    return VisitExtendedNode(evalNode);
                
                default:
                    return Array.Empty<CodePath>();
            }
        }

        private void HitNode(IExtendedNode node)
        {
            node.IsReachable = true;
            _currentPath.Peek().Add(node);
            _position++; // note: a jump node can change this.
        }

        private CodePath[] VisitExtendedNode(IEvaluatableNode node)
        {
            return Array.Empty<CodePath>();
        }

        private CodePath[] VisitExtendedNode(IExecutableNode node)
        {
            var paths = new List<CodePath>();
            if (node is VBAParser.ExitStmtContext exit && exit.ExitsScope)
            {
                throw new Exception("ExitStmtContext.ExitsScope"); // todo: don't throw anything
            }

            var body = ((IParseTree)node).FlattenExtendedNodes();
            foreach (var item in body)
            {
                paths.AddRange(VisitExtendedNode(item));
            }

            return paths.ToArray();
        }

        private CodePath[] VisitExtendedNode(IBranchNode node)
        {
            HitNode(node.ConditionExpression);
            var paths = new List<CodePath>();
            EnterExecutionPath();
            
            var body = ((IParseTree)node).FlattenExtendedNodes().Skip(1); // skip ConditionExpression
            foreach (var item in body)
            {
                paths.AddRange(VisitExtendedNode(item));
            }

            ExitExecutionPath();
            return paths.ToArray();
        }

        public static readonly int MaxPathIterations = 8;

        private readonly IDictionary<IJumpNode, int> _jumps
            = new Dictionary<IJumpNode, int>();

        private CodePath[] VisitExtendedNode(IJumpNode node)
        {            
            if (_jumps.ContainsKey(node))
            {
                if (_jumps[node] >= MaxPathIterations)
                {
                    throw new MaxIterationsReachedException(node);
                }
                _jumps[node] = _jumps[node]++;
            }
            else
            {
                _jumps.Add(node, 1);
            }
            _position = Array.IndexOf(_nodes, node.Target);
            return Array.Empty<CodePath>();
        }

        private void EnterExecutionPath()
        {
            var path = new CodePath();
            _allPaths.Add(path);
            _currentPath.Push(path);
        }

        private void ExitExecutionPath()
        {
            _currentPath.Pop();
        }
    }
}
