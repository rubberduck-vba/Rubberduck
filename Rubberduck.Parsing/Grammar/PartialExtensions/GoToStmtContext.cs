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
        public partial class GoToStmtContext : IJumpNode
        {
            public IExtendedNode Target { get; set; }

            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
                => _hasExecuted[context] = true;
        }

        public partial class OnErrorStmtContext : IJumpNode
        {
            public IExtendedNode Target { get; set; }

            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
                => _hasExecuted[context] = true;
        }

        public partial class ResumeStmtContext : IJumpNode
        {
            public IExtendedNode Target { get; set; }

            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
                => _hasExecuted[context] = true;
        }

        public partial class ReturnStmtContext : IJumpNode
        {
            public IExtendedNode Target { get; set; }

            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
                => _hasExecuted[context] = true;
        }

        public partial class GoSubStmtContext : IJumpReferenceNode
        {
            public IExtendedNode Target { get; set; }

            private readonly IDictionary<IExecutionContext, bool> _hasExecuted 
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
                => _hasExecuted[context] = true;


            private readonly IDictionary<IExecutionContext, IJumpNode> _origin
                = new Dictionary<IExecutionContext, IJumpNode>();

            public void SetOrigin(IJumpNode node, IExecutionContext context) 
                => _origin[context] = node;

            public IJumpNode GetOrigin(IExecutionContext context) 
                => _origin.TryGetValue(context, out var node) ? node : null;
        }
    }
}
