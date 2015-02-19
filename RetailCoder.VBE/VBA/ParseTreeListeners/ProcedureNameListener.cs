using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            if (context.AmbiguousIdentifier().GetText() == _name)
            {
                base.EnterFunctionStmt(context);
            }
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            if (context.AmbiguousIdentifier().GetText() == _name)
            {
                base.EnterSubStmt(context);
            }
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            if (context.AmbiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertyGetStmt(context);
            }
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            if (context.AmbiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertyLetStmt(context);
            }
        }

        public override void EnterPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            if (context.AmbiguousIdentifier().GetText() == _name)
            {
                base.EnterPropertySetStmt(context);
            }
        }
    }
}
