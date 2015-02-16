using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : VisualBasic6BaseListener, IExtensionListener<VisualBasic6Parser.AmbiguousIdentifierContext>
    {
        private readonly IList<VisualBasic6Parser.AmbiguousIdentifierContext> _members = new List<VisualBasic6Parser.AmbiguousIdentifierContext>();
        public IEnumerable<VisualBasic6Parser.AmbiguousIdentifierContext> Members { get { return _members; } }

        public override void EnterForNextStmt(VisualBasic6Parser.ForNextStmtContext context)
        {
            _members.Add(context.ambiguousIdentifier().First());
        }

        public override void EnterVariableCallStmt(VisualBasic6Parser.VariableCallStmtContext context)
        {
            if (context.Parent.Parent.Parent is VisualBasic6Parser.LetStmtContext)
            {
                _members.Add(context.ambiguousIdentifier());
            }
        }
    }
}