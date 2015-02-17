using System.Collections.Generic;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class PublicSubListener : IVBBaseListener, IExtensionListener<VBParser.SubStmtContext>
    {
        private readonly IList<VBParser.SubStmtContext> _members = new List<VBParser.SubStmtContext>();
        public IEnumerable<VBParser.SubStmtContext> Members { get { return _members; } }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            var visibility = context.visibility();
            if (visibility == null || visibility.PUBLIC() != null)
            {
                _members.Add(context);
            }
        }
    }
}