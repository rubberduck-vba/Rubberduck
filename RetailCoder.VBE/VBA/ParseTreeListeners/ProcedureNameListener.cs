using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ProcedureNameListener : ProcedureListener
    {
        private readonly string _name;

        public ProcedureNameListener(string name)
        {
            _name = name;
        }

        public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterFunctionStmt(context);
            }
        }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterSubStmt(context);
            }
        }

        public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertyGetStmt(context);
            }
        }

        public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertyLetStmt(context);
            }
        }

        public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertySetStmt(context);
            }
        }
    }
}