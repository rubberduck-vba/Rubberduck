using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : VariableUsageListener
    {
        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            if (!(context.Parent is VBParser.AsTypeClauseContext))
            {
                base.EnterVariableCallStmt(context);
            }
        }

        public override void EnterFunctionOrArrayCallStmt(VBParser.FunctionOrArrayCallStmtContext context)
        {
            AddMember(context.AmbiguousIdentifier());
        }

        public override void EnterVariableSubStmt(VBParser.VariableSubStmtContext context)
        {
            // consider "As New [classname]" as an assignemnt:
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