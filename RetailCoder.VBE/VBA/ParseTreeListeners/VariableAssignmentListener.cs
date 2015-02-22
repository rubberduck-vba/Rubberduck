using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableAssignmentListener : VariableUsageListener
    {
        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            if (context.Parent.Parent.Parent is VBParser.LetStmtContext)
            {
                base.EnterVariableCallStmt(context);
            }
        }

        public VariableAssignmentListener(QualifiedModuleName qualifiedName) 
            : base(qualifiedName)
        {
        }
    }
}