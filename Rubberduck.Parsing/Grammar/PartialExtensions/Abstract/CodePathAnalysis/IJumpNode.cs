namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node representing a jump (up or down) in execution.
    /// </summary>
    public interface IJumpNode : IExecutableNode
    {
        /// <summary>
        /// Gets or sets the jump target.
        /// </summary>
        IExtendedNode Target { get; set; }
    }
}