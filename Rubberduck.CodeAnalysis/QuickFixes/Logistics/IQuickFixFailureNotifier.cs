using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Logistics
{
    internal interface IQuickFixFailureNotifier
    {
        void NotifyQuickFixExecutionFailure(RewriteSessionState sessionState);
    }
}