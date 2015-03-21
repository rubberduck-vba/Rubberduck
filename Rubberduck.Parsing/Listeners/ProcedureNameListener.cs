using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Listeners
{
    public class ProcedureNameListener : ProcedureListener
    {
        private readonly string _name;

        public ProcedureNameListener(string name, QualifiedModuleName qualifiedName)
            : base(qualifiedName)
        {
            _name = name;
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterFunctionStmt(context);
            }
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterSubStmt(context);
            }
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertyGetStmt(context);
            }
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertyLetStmt(context);
            }
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            if (context.ambiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertySetStmt(context);
            }
        }
    }
}
