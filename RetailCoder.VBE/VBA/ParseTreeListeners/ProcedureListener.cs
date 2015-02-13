using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ProcedureListener : VisualBasic6BaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
        public IEnumerable<ParserRuleContext> Members { get { return _members; } }

        public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
        {
            _members.Add(context);
        }
    }
}