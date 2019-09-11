namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node that can be executed using a specified <see cref="IExecutionContext"/>.
    /// </summary>
    public interface IExecutableNode : IExtendedNode
    {
        /// <summary>
        /// <c>true</c> if the node has executed in the specified <see cref="context"/>.
        /// </summary>
        bool HasExecuted(IExecutionContext context);
        /// <summary>
        /// Executes the node for the specified <see cref="context"/>.
        /// </summary>
        void Execute(IExecutionContext context);
    }
}