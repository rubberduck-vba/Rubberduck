namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node that represents a reference to a <see cref="Declaration"/>.
    /// </summary>
    public interface IReferenceNode : IExtendedNode
    {
        /// <summary>
        /// Marks the reference as assigned in the specified context.
        /// </summary>
        void Assign(IExecutionContext context);
        /// <summary>
        /// Gets a value indicating whether the reference is assigned in the specified context.
        /// </summary>
        /// <remarks>Given an object reference, an assignment to <c>Nothing</c> should not make this method return <c>true</c>.</remarks>
        bool IsAssigned(IExecutionContext context);
    }
}