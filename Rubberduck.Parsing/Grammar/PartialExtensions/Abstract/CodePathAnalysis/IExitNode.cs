namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// A node representing a jump outside of a parent loop block or procedure scope.
    /// </summary>
    public interface IExitNode : IExecutableNode
    { 
        bool ExitsScope { get; }
    }
}