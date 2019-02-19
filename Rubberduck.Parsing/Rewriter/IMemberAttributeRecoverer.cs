using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface IMemberAttributeRecoverer
    {
        void RecoverCurrentMemberAttributesAfterNextParse(IEnumerable<QualifiedMemberName> members);
        void RecoverCurrentMemberAttributesAfterNextParse(IEnumerable<QualifiedModuleName> modules);
    }
}