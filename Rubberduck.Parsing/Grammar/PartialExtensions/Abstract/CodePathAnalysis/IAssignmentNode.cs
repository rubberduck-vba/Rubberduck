using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node that represents an assignment operation and tracks references to its value.
    /// </summary>
    public interface IAssignmentNode : IExecutableNode
    {
        /// <summary>
        /// Gets all references to this assignment in the specified <see cref="IExecutionContext"/>.
        /// </summary>
        IImmutableSet<IReferenceNode> References(IExecutionContext context);
        /// <summary>
        /// Adds a reference to the value of this assignment operation in the specified <see cref="IExecutionContext"/>.
        /// </summary>
        void AddReference(IReferenceNode node, IExecutionContext context);
    }
}
