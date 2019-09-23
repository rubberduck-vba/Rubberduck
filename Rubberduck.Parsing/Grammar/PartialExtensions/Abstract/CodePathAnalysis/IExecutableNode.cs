namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node that can be executed using a specified <see cref="IExecutionContext"/>.
    /// </summary>
    public interface IExecutableNode : IExtendedNode
    {
        /// <summary>
        /// Executes the node for the specified <see cref="context"/>.
        /// </summary>
        void Execute(IExecutionContext context);
    }
}