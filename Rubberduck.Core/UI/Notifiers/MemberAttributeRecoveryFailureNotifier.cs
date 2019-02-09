using System;
using System.Collections.Generic;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Notifiers
{
    public class MemberAttributeRecoveryFailureNotifier : IMemberAttributeRecoveryFailureNotifier
    {
        private readonly IMessageBox _messageBox;

        public MemberAttributeRecoveryFailureNotifier(IMessageBox messageBox)
        {
            _messageBox = messageBox;
        }

        public void NotifyRewriteFailed(RewriteSessionState sessionState)
        {
            var message = RewriteFailureMessage(sessionState);
            var caption = Resources.RubberduckUI.MemberAttributeRecoveryFailureCaption;

            _messageBox.NotifyWarn(message, caption);
        }

        private static string RewriteFailureMessage(RewriteSessionState sessionState)
        {
            var baseFailureMessage = Resources.RubberduckUI.MemberAttributeRecoveryRewriteFailedMessage;
            var failureReasonMessage = RewriteFailureReasonMessage(sessionState);
            var message = string.IsNullOrEmpty(failureReasonMessage)
                ? baseFailureMessage
                : $"{baseFailureMessage}{Environment.NewLine}{Environment.NewLine}{failureReasonMessage}";
            return message;
        }

        private static string RewriteFailureReasonMessage(RewriteSessionState sessionState)
        {
            switch (sessionState)
            {
                case RewriteSessionState.StaleParseTree:
                    return Resources.Inspections.QuickFixes.StaleModuleFailureReason;
                default:
                    return string.Empty;
            }
        }

        public void NotifyMembersForRecoveryNotFound(IEnumerable<QualifiedMemberName> membersNotFound)
        {
            var message = MembersNotFoundMessage(membersNotFound);
            var caption = Resources.RubberduckUI.MemberAttributeRecoveryFailureCaption;

            _messageBox.NotifyWarn(message, caption);
        }

        private string MembersNotFoundMessage(IEnumerable<QualifiedMemberName> membersNotFound)
        {
            var missingMemberList = $"{Environment.NewLine}{string.Join(Environment.NewLine, membersNotFound)}";
            return string.Format(Resources.RubberduckUI.MemberAttributeRecoveryMembersNotFoundMessage, missingMemberList); ;
        }
    }
}