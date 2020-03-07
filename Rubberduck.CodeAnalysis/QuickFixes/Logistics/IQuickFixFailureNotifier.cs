using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    internal interface IQuickFixFailureNotifier
    {
        void NotifyQuickFixExecutionFailure(RewriteSessionState sessionState);
    }
}