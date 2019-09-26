namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// An executable node that contains an <see cref="IEvaluatableNode"/> that determines if a child block is entered.
    /// </summary>
    /// <remarks>
    /// The <see cref="IExecutableNode.HasExecuted"/> state of this node indicates
    /// whether the <see cref="ConditionExpression"/> was evaluated.
    /// </remarks>
    public interface IBranchNode : IExecutableNode
    {
        /// <summary>
        /// Gets the <see cref="IEvaluatableNode"/> that contains the conditional expression for branching.
        /// </summary>
        IEvaluatableNode ConditionExpression { get; }
    }
}