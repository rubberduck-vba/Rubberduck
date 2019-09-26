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
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing;
using Antlr4.Runtime;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution.ExtendedNodeVisitor
{
    public class ExtendedNodeVisitor
    {
        private readonly IExtendedNode[] _nodes;
        private readonly HashSet<ILabelNode> _labels;
        private readonly HashSet<IdentifierReference> _refs;

        public ExtendedNodeVisitor(ModuleBodyElementDeclaration member, IDeclarationFinderProvider provider)
        {
            var finder = provider.DeclarationFinder;
            _refs = new HashSet<IdentifierReference>(finder.IdentifierReferences(member.QualifiedName));

            _nodes = member.Context.FlattenExtendedNodes().ToArray();
            _labels = new HashSet<ILabelNode>(_nodes.OfType<ILabelNode>());
        }

        private readonly HashSet<CodePath> _allPaths = new HashSet<CodePath>();
        private readonly Stack<CodePath> _currentPath = new Stack<CodePath>();
        private readonly MergedPath _mergedPath = new MergedPath();

        private int _position = 0;

        private void EnterExecutionPath()
        {
            var current = _currentPath.Any() ? _currentPath.Peek() : null;
            var path = current?.Clone() ?? new CodePath();
            _allPaths.Add(path);
            _currentPath.Push(path);
        }

        private void ExitExecutionPath()
        {
            var path = _currentPath.Pop();
            _mergedPath.Merge(path);
        }

        public CodePath[] GetAllCodePaths()
        {
            var traversed = new List<IExtendedNode>(); // todo: remove

            _position = 0;
            EnterExecutionPath();

            while (_position < _nodes.Length)
            {
                var node = _nodes[_position];

                #region safety net code to be deleted
                if (node == traversed.LastOrDefault())
                {
                    // which is more likely:
                    // a self-referencing goto-loop, or a bug somewhere?
                    Debug.Assert(false);
                }
                traversed.Add(node);
                #endregion
                
                var paths = VisitExtendedNode(node);
                foreach (var path in paths)
                {
                    _allPaths.Add(path);
                }
            }

            return _allPaths.ToArray();
        }

        private CodePath[] VisitExtendedNode(IExtendedNode node)
        {
            switch(node)
            {
                case IAssignmentNode assignment:
                    if (assignment.Target == null)
                    {
                        assignment.Target = _refs
                            .SingleOrDefault(r => r.IsAssignment && r.Context.GetSelection().IsContainedIn(assignment as ParserRuleContext));
                    }
                    return VisitExtendedNode(assignment);
                case IExitNode exitNode:
                    return VisitExtendedNode(exitNode);
                case IBranchNode branchNode:
                    return VisitExtendedNode(branchNode);
                case IJumpNode jumpNode:
                    return VisitExtendedNode(jumpNode);
                case IEvaluatableNode evalNode:
                    return VisitExtendedNode(evalNode);
                case IExecutableNode exeNode:
                    return VisitExtendedNode(exeNode);
                
                default:
                    return Array.Empty<CodePath>();
            }
        }

        private void HitNode(IExtendedNode node)
        {
            if (node == null)
            {
                return;
            }

            node.IsReachable = true;
            var path = _currentPath.Peek();
            path.Add(node);
            _mergedPath.Add(node);
            _position++; // note: a jump node can change this.
        }

        private CodePath[] VisitExtendedNode(IExecutableNode node)
        {
            HitNode(node);
            var paths = new List<CodePath>();
            if (node is IExitNode exit && exit.ExitsScope)
            {
                throw new Exception("ExitStmtContext.ExitsScope"); // todo: don't throw anything... somehow
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
            HitNode(node);
            if (node.ConditionExpression != null)
            {
                // node is evaluated in current path
                HitNode(node.ConditionExpression); 
            }
            var paths = new List<CodePath>();
            EnterExecutionPath();
            
            var body = ((IParseTree)node).FlattenExtendedNodes()
                .Where(n => typeof(IExecutableNode).IsAssignableFrom(n.GetType()));

            foreach (var item in body)
            {
                if (item is IBranchNode branch && branch.ConditionExpression == null) 
                {
                    // else block must exit current path
                    break; 
                }
                paths.AddRange(VisitExtendedNode(item));                
            }

            ExitExecutionPath();
            return paths.ToArray();
        }

        public static readonly int MaxPathIterations = 8;

        private CodePath[] VisitExtendedNode(IJumpNode node)
        {
            HitNode(node);
            var jumps = new Dictionary<IJumpNode, int>();
            if (jumps.ContainsKey(node))
            {
                if (jumps[node] >= MaxPathIterations)
                {
                    throw new MaxIterationsReachedException(node);
                }
                jumps[node] = jumps[node]++;
            }
            else
            {
                jumps.Add(node, 1);
            }
            _position = Array.IndexOf(_nodes, node.Target);
            return Array.Empty<CodePath>();
        }

        private CodePath[] VisitExtendedNode(IEvaluatableNode node)
        {
            HitNode(node);
            var refs = _refs.Where(r => ((ParserRuleContext)node).ContainsTokenIndex(r.Context.Start.TokenIndex));
            foreach (var identifierRef in refs)
            {
                var path = _currentPath.Peek();
                if (identifierRef.IsAssignment)
                {
                    if (((ParserRuleContext)node).TryGetAncestor<IAssignmentNode>(out var assignment))
                    {
                        path.OnAssignment(identifierRef, assignment);
                    }
                    else
                    {
                        Debug.Assert(false, "Assignment node ancestor not found.");
                    }
                }
                else
                {
                    path.OnReference(identifierRef, node);
                }
            }
            return Array.Empty<CodePath>();
        }
    }
}
