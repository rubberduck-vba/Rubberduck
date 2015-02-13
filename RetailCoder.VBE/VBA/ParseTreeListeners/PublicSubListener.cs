using System.Collections.Generic;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class PublicSubListener : VisualBasic6BaseListener, IExtensionListener<VisualBasic6Parser.SubStmtContext>
    {
        private readonly IList<VisualBasic6Parser.SubStmtContext> _members = new List<VisualBasic6Parser.SubStmtContext>();
        public IEnumerable<VisualBasic6Parser.SubStmtContext> Members { get { return _members; } }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            var visibility = context.visibility();
            if (visibility == null || visibility.PUBLIC() != null)
            {
                _members.Add(context);
            }
        }
    }
}