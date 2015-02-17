using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class DeclarationSectionListener : DeclarationListener
    {
        public override void EnterArg(VisualBasic6Parser.ArgContext context)
        {
            return;
        }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void ExitPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            throw new WalkerCancelledException();
        }
    }
}