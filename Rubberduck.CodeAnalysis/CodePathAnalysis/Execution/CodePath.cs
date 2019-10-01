using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution
{
    public class CodePath
    {
        private readonly List<IExtendedNode> _nodes;

        private readonly Dictionary<IdentifierReference, Stack<IAssignmentNode>> _refs
            = new Dictionary<IdentifierReference, Stack<IAssignmentNode>>();
        
        private readonly Dictionary<IAssignmentNode, Stack<IEvaluatableNode>> _assignmentReads
            = new Dictionary<IAssignmentNode, Stack<IEvaluatableNode>>();

        private readonly HashSet<IAssignmentNode> _seededRefs 
            = new HashSet<IAssignmentNode>();

        public CodePath(IEnumerable<IExtendedNode> nodes = null, bool isErrorPath = false)
        {
            _nodes = new List<IExtendedNode>(nodes ?? Enumerable.Empty<IExtendedNode>());
            IsErrorPath = isErrorPath;
        }

        internal void SeedRefs(Dictionary<IdentifierReference, Stack<IAssignmentNode>> refs)
        {
            foreach (var kvp in refs)
            {
                if (!_refs.ContainsKey(kvp.Key) || _refs[kvp.Key] == null)
                {
                    _refs.Add(kvp.Key, new Stack<IAssignmentNode>());
                }
                var lastRef = kvp.Value.Peek();
                if (lastRef != null)
                {
                    _refs[kvp.Key].Push(lastRef);
                    _seededRefs.Add(lastRef);
                }
            }
        }

        internal CodePath Clone()
        {
            var result = new CodePath(isErrorPath:this.IsErrorPath);
            result.SeedRefs(_refs);
            return result;
        }

        public bool IsErrorPath { get; }

        public IExtendedNode this[int index] 
            => _nodes[index];

        public IExtendedNode[] Nodes(int index) 
            => _nodes.ToArray();

        public int Count 
            => _nodes.Count;

        /// <summary>
        /// Gets all assignment nodes of this path, plus the last assignment node of the previous path.
        /// No assignment nodes for a reference means there is no assignment.
        /// Use <see cref="Assignments"/> to get the assignment nodes for this path.
        /// </summary>
        internal IDictionary<IdentifierReference, Stack<IAssignmentNode>> AssignmentMetadata 
            => _refs;

        /// <summary>
        /// Gets all assignment nodes of this path.
        /// No assignment nodes for a reference means it isn't assigned *in this code path*.
        /// </summary>
        internal IDictionary<IdentifierReference, IAssignmentNode[]> Assignments
            => _refs.Select(kvp => new 
                {
                    kvp.Key, 
                    Value = kvp.Value.Where(v => !_seededRefs.Contains(v)).ToArray()
                })
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

        internal void OnAssignment(IdentifierReference reference, IAssignmentNode node = null)
        {
            Debug.Assert(reference.IsAssignment);

            if (!_refs.ContainsKey(reference))
            {
                _refs.Add(reference, new Stack<IAssignmentNode>());
            }

            if (_refs[reference] == null)
            {
                _refs[reference] = new Stack<IAssignmentNode>();
            }

            var assignment = node ?? LastAssignment(reference);

            _refs[reference].Push(assignment);
            _assignmentReads[assignment] = new Stack<IEvaluatableNode>();
        }

        internal void OnReference(IdentifierReference reference, IEvaluatableNode node)
        {
            if (reference.IsAssignment)
            {
                OnAssignment(reference);
            }
            else
            {
                var assignment = LastAssignment(reference);
                if (assignment != null && _assignmentReads.TryGetValue(assignment, out _))
                {
                    _assignmentReads[assignment].Push(node);
                }
                else if (assignment != null)
                {
                    // todo: flag unassigned read?
                }
            }
        }

        internal IAssignmentNode LastAssignment(IdentifierReference reference)
        {
            if (_refs.TryGetValue(reference, out var stack))
            {
                return (stack?.Any() ?? false) ? stack.Peek() : null;
            }
            return null;
        }

        internal void Add(IExtendedNode node) 
            => _nodes.Add(node);

        internal void AddRange(IEnumerable<IExtendedNode> nodes)
        {
            foreach (var node in nodes)
            {
                Add(node);
            }
        }

        public IAssignmentNode[] UnreferencedAssignments 
            => _assignmentReads.Where(kvp => !kvp.Value.Any()).Select(kvp => kvp.Key).ToArray();

        public IExtendedNode[] UnreachableNodes 
            => _nodes.Where(node => !node.IsReachable).ToArray();

        public override string ToString() 
            => string.Join("\n", _nodes.Select((node, index) => $"[{index}]:{node.GetType().Name}"));
    }

    internal class MergedPath
    {
        private readonly List<CodePath> _paths = new List<CodePath>();

        public void Merge(CodePath path) => _paths.Add(path);

        public void Add(IExtendedNode node) => _paths.ForEach(path => path.Add(node));

        public CodePath[] Flatten() => _paths.ToArray();
    }
}
