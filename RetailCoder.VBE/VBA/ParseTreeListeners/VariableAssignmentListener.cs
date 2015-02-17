using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : IVBBaseListener, IExtensionListener<VBParser.AmbiguousIdentifierContext>
    {
        private readonly IList<VBParser.AmbiguousIdentifierContext> _members = new List<VBParser.AmbiguousIdentifierContext>();
        public IEnumerable<VBParser.AmbiguousIdentifierContext> Members { get { return _members; } }

        public override void EnterForNextStmt(VBParser.ForNextStmtContext context)
        {
            _members.Add(context.ambiguousIdentifier().First());
        }

        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            if (context.Parent.Parent.Parent is VBParser.LetStmtContext)
            {
                _members.Add(context.ambiguousIdentifier());
            }
        }
    }
}