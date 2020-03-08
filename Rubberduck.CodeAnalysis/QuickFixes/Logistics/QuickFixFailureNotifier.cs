using System;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Logistics
{
    internal class QuickFixFailureNotifier : IQuickFixFailureNotifier
    {
        private readonly IMessageBox _messageBox;

        public QuickFixFailureNotifier(IMessageBox messageBox)
        {
            _messageBox = messageBox;
        }

        public void NotifyQuickFixExecutionFailure(RewriteSessionState sessionState)
        {
            var message = FailureMessage(sessionState);
            var caption = Resources.Inspections.QuickFixes.ApplyQuickFixFailedCaption;

            _messageBox.NotifyWarn(message, caption);
        }

        private static string FailureMessage(RewriteSessionState sessionState)
        {
            var baseFailureMessage = Resources.Inspections.QuickFixes.ApplyQuickFixesFailedMessage;
            var failureReasonMessage = FailureReasonMessage(sessionState);
            var message = string.IsNullOrEmpty(failureReasonMessage)
                ? baseFailureMessage
                : $"{baseFailureMessage}{Environment.NewLine}{Environment.NewLine}{failureReasonMessage}";
            return message;
        }

        private static string FailureReasonMessage(RewriteSessionState sessionState)
        {
            switch (sessionState)
            {
                case RewriteSessionState.StaleParseTree:
                    return Resources.Inspections.QuickFixes.StaleModuleFailureReason;
                default:
                    return string.Empty;
            }
        }
    }
}