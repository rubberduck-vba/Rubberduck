using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Refactorings.Exceptions
{
    public class RewriteFailedException : RefactoringException
    {
        public RewriteFailedException(IRewriteSession rewriteSession)
        {
            RewriteSession = rewriteSession;
        }

        public IRewriteSession RewriteSession { get; }
    }
}