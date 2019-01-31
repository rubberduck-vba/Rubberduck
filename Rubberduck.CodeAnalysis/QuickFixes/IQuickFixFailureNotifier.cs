using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public interface IQuickFixFailureNotifier
    {
        void NotifyQuickFixExecutionFailure(RewriteSessionState sessionState);
    }
}