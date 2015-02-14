using System.Collections.Generic;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : VisualBasic6BaseListener, IExtensionListener<VisualBasic6Parser.VariableCallStmtContext>
    {
        private readonly IList<VisualBasic6Parser.VariableCallStmtContext> _members = new List<VisualBasic6Parser.VariableCallStmtContext>();
        public IEnumerable<VisualBasic6Parser.VariableCallStmtContext> Members { get { return _members; } }

        public override void EnterVariableCallStmt(VisualBasic6Parser.VariableCallStmtContext context)
        {
            if (context.Parent.Parent.Parent is VisualBasic6Parser.LetStmtContext)
            {
                _members.Add(context);
            }
        }
    }
}