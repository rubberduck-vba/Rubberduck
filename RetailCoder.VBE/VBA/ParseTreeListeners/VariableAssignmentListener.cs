using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : VariableUsageListener
    {
        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            if ((context.Parent.Parent.Parent is VBParser.LetStmtContext
                || context.Parent.Parent.Parent is VBParser.SetStmtContext
                || context.Parent.Parent.Parent.Parent is VBParser.ValueStmtContext)
            && !(context.Parent is VBParser.AsTypeClauseContext))
            {
                base.EnterVariableCallStmt(context);
            }
        }

        public override void EnterVariableSubStmt(VBParser.VariableSubStmtContext context)
        {
            if (context.AsTypeClause() != null && context.AsTypeClause().NEW() != null)
            {
                AddMember(context.AmbiguousIdentifier());
            }
        }

        public VariableAssignmentListener(QualifiedModuleName qualifiedName) 
            : base(qualifiedName)
        {
        }
    }
}