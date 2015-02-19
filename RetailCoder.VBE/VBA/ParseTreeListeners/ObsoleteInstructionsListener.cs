using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ObsoleteInstrutionsListener : VBListenerBase, IExtensionListener<ParserRuleContext>
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

        public override void EnterLetStmt(VBParser.LetStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterExplicitCallStmt(VBParser.ExplicitCallStmtContext context)
        {

            if (context.ECS_MemberProcedureCall() != null 
                && context.ECS_MemberProcedureCall().CALL() != null)
            {
                AddMember(context.ECS_MemberProcedureCall());
            }
            else if (context.ECS_ProcedureCall() != null
                && context.ECS_ProcedureCall().CALL() != null)
            {
                AddMember(context.ECS_ProcedureCall());
            }
        }
    }
}
