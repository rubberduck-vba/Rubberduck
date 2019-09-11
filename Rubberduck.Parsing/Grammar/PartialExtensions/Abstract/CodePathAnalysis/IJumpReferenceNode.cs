namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node representing a jump that retains its origin, like a <c>GoSubStmt</c>.
    /// </summary>
    public interface IJumpReferenceNode : IJumpNode
    {
        /// <summary>
        /// Sets the jump origin in the specified execution context.
        /// </summary>
        void SetOrigin(IJumpNode node, IExecutionContext context);
        /// <summary>
        /// Gets the jump origin in the specified execution context.
        /// </summary>
        IJumpNode GetOrigin(IExecutionContext context);
    }
}