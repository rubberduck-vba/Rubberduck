using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IMemberAttributeRecoveryFailureNotifier
    {
        void NotifyRewriteFailed(RewriteSessionState rewriteSessionState);
        void NotifyMembersForRecoveryNotFound(IEnumerable<(QualifiedMemberName memberName, DeclarationType memberType)> membersNotFound);
    }
}