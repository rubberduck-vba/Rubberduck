using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution
{
    /// <summary>
    /// Represents an assignable value that is accessible in the current procedure scope.
    /// </summary>
    public struct AssignableValue
    {
        /// <summary>
        /// Wraps a declaration so that it can be tied to an assignment node.
        /// </summary>
        public AssignableValue(Declaration declaration) 
            : this(declaration, null) { }

        private AssignableValue(Declaration declaration, IAssignmentNode value)
        {
            Declaration = declaration;
            Value = value;
        }

        public Declaration Declaration { get; }

        /// <summary>
        /// True if <see cref="Value"/> isn't <c>null</c>.
        /// </summary>
        public bool IsAssigned => !(Value is object);

        /// <summary>
        /// Gets the node that assigned the current value.
        /// </summary>
        public IAssignmentNode Value { get; }

        /// <summary>
        /// Gets a new <see cref="AssignableValue"/> pointing to the specified assignment operation.
        /// </summary>
        public AssignableValue Assign(IAssignmentNode node) 
            => new AssignableValue(Declaration, node);
    }
}
