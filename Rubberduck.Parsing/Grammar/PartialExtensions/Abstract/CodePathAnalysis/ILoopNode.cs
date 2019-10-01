namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node representing a loop control structure.
    /// </summary>
    public interface ILoopNode : IExecutableNode
    {
        /// <summary>
        /// Gets or sets the conditional expression that determines whether the loop is entered or exited.
        /// </summary>
        IEvaluatableNode ConditionExpression { get; }
    }
}