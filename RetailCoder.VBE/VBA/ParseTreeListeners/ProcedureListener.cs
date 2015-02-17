using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ProcedureListener : IVBBaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
        public IEnumerable<ParserRuleContext> Members { get { return _members; } }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            _members.Add(context);
        }
    }
}