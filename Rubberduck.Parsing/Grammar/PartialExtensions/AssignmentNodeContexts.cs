using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class SetStmtContext : IAssignmentNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            private readonly IDictionary<IExecutionContext, IList<IReferenceNode>> _references
                = new Dictionary<IExecutionContext, IList<IReferenceNode>>();

            public void AddReference(IReferenceNode node, IExecutionContext context) 
            {
                if (!_references.TryGetValue(context, out var refs) || refs is null)
                {
                    _references[context] = refs = new List<IReferenceNode>();
                }
                refs.Add(node);
            }

            public IReadOnlyList<IReferenceNode> References(IExecutionContext context) 
            {
                if (!_references.TryGetValue(context, out var refs) || refs is null)
                {
                    _references[context] = refs = new List<IReferenceNode>();
                }
                return refs.ToList();
            }
        }

        public partial class LetStmtContext : IAssignmentNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            private readonly IDictionary<IExecutionContext, IList<IReferenceNode>> _references
                = new Dictionary<IExecutionContext, IList<IReferenceNode>>();

            public void AddReference(IReferenceNode node, IExecutionContext context) 
            {
                if (!_references.TryGetValue(context, out var refs) || refs is null)
                {
                    _references[context] = refs = new List<IReferenceNode>();
                }
                refs.Add(node);
            }

            public IReadOnlyList<IReferenceNode> References(IExecutionContext context) 
            {
                if (!_references.TryGetValue(context, out var refs) || refs is null)
                {
                    _references[context] = refs = new List<IReferenceNode>();
                }
                return refs.ToList();
            }
        }
    }
}
