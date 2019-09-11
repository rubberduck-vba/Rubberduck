namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node representing an evaluatable expression that can resolve to a specified data type.
    /// </summary>
    public interface IEvaluatableNode : IExtendedNode
    {
        /// <summary>
        /// Evaluates the expression in the specified execution context, and returns a result of the specified type.
        /// </summary>
        T Evaluate<T>(IExecutionContext context);
    }
}