namespace Rubberduck.Parsing.Rewriter
{
    public interface IExecutableModuleRewriter : IModuleRewriter
    {
        /// <summary>
        /// Rewrites the entire module / applies all changes.
        /// </summary>
        void Rewrite();
    }
}