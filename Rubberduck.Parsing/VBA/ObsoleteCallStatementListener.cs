using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA
{
    public class ObsoleteCallStatementListener : VBABaseListener
    {
        private readonly IList<VBAParser.ExplicitCallStmtContext> _contexts = new List<VBAParser.ExplicitCallStmtContext>();
        public IEnumerable<VBAParser.ExplicitCallStmtContext> Contexts { get { return _contexts; } }

        public override void EnterExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
        {
            var procedureCall = context.eCS_ProcedureCall();
            if (procedureCall != null)
            {
                if (procedureCall.CALL() != null)
                {
                    _contexts.Add(context);
                    return;
                }
            }

            var memberCall = context.eCS_MemberProcedureCall();
            if (memberCall == null) return;
            if (memberCall.CALL() == null) return;
            _contexts.Add(context);
        }
    }
}