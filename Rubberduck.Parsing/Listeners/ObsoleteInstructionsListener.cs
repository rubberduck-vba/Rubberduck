using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Listeners
{
    public class ObsoleteInstrutionsListener : VBABaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<ParserRuleContext>> _members = 
            new List<QualifiedContext<ParserRuleContext>>();

        public IEnumerable<QualifiedContext<ParserRuleContext>> Members { get { return _members; } }

        public ObsoleteInstrutionsListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        private void AddMember<TContext>(TContext context) where TContext : ParserRuleContext
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterLetStmt(VBAParser.LetStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterExplicitCallStmt(VBAParser.ExplicitCallStmtContext context)
        {

            if (context.eCS_MemberProcedureCall() != null 
                && context.eCS_MemberProcedureCall().CALL() != null)
            {
                AddMember(context.eCS_MemberProcedureCall());
            }
            else if (context.eCS_ProcedureCall() != null
                && context.eCS_ProcedureCall().CALL() != null)
            {
                AddMember(context.eCS_ProcedureCall());
            }
        }
    }
}
