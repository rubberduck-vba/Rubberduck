using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

            if (context.ECS_MemberProcedureCall() != null)
            {
                if (context.ECS_MemberProcedureCall().CALL() != null)
                {
                    _members.Add(context.ECS_MemberProcedureCall());
                }
            }
            else if (context.ECS_ProcedureCall() != null)
            {
                if (context.ECS_ProcedureCall().CALL() != null)
                {
                    _members.Add(context.ECS_ProcedureCall());
                }
            }
        }
    }
}
