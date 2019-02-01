using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IMemberAttributeRecoveryFailureNotifier
    {
        void NotifyRewriteFailed(RewriteSessionState rewriteSessionState);
        void NotifyMembersForRecoveryNotFound(IEnumerable<QualifiedMemberName> membersNotFound);
    }
}