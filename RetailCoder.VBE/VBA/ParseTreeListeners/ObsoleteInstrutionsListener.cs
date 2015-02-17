using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ObsoleteInstrutionsListener : IVBBaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
        public IEnumerable<ParserRuleContext> Members { get { return _members; } }

        public override void EnterLetStmt(VBParser.LetStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterExplicitCallStmt(VBParser.ExplicitCallStmtContext context)
        {

            if (context.eCS_MemberProcedureCall() != null)
            {
                if (context.eCS_MemberProcedureCall().CALL() != null)
                {
                    _members.Add(context.eCS_MemberProcedureCall());
                }
            }
            else if (context.eCS_ProcedureCall() != null)
            {
                if (context.eCS_ProcedureCall().CALL() != null)
                {
                    _members.Add(context.eCS_ProcedureCall());
                }
            }
        }
    }
}