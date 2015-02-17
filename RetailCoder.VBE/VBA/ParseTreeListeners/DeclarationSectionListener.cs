using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class DeclarationListener : IVBBaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
        public IEnumerable<ParserRuleContext> Members { get { return _members; } }

        public override void EnterVariableStmt(VBParser.VariableStmtContext context)
        {
            _members.Add(context);
            foreach (var child in context.variableListStmt().variableSubStmt())
            {
                _members.Add(child);
            }
        }

        public override void EnterEnumerationStmt(VBParser.EnumerationStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterConstStmt(VBParser.ConstStmtContext context)
        {
            _members.Add(context);
            foreach (var child in context.constSubStmt())
            {
                _members.Add(child);
            }
        }

        public override void EnterTypeStmt(VBParser.TypeStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterDeclareStmt(VBParser.DeclareStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterEventStmt(VBParser.EventStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterArg(VBParser.ArgContext context)
        {
            _members.Add(context);
        }
    }

    public class DeclarationSectionListener : DeclarationListener
    {
        public override void EnterArg(VBParser.ArgContext context)
        {
            return;
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void ExitPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            throw new WalkerCancelledException();
        }
    }
}