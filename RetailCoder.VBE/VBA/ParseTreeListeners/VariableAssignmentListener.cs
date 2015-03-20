using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.Parsing;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : VariableUsageListener
    {
        public override void EnterICS_S_VariableOrProcedureCall(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            if (context.Parent is VBAParser.ImplicitCallStmt_InStmtContext
                && (context.Parent.Parent is VBAParser.LetStmtContext
                    || context.Parent.Parent is VBAParser.SetStmtContext))
            {
                base.EnterICS_S_VariableOrProcedureCall(context);
            }
        }

        public override void EnterForNextStmt(VBAParser.ForNextStmtContext context)
        {
            AddMember(context.ambiguousIdentifier().First());
        }

        public override void EnterForEachStmt(VBAParser.ForEachStmtContext context)
        {
            AddMember(context.ambiguousIdentifier().First());
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            // consider "As New [classname]" as an assignemnt:
            if (context.asTypeClause() != null && context.asTypeClause().NEW() != null)
            {
                AddMember(context.ambiguousIdentifier());
            }
        }

        public VariableAssignmentListener(QualifiedModuleName qualifiedName) 
            : base(qualifiedName)
        {
        }
    }
}