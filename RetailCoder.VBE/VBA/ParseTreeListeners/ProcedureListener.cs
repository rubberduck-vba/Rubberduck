using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ProcedureListener : VBListenerBase, IExtensionListener<ParserRuleContext>
    {
        protected readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<ParserRuleContext>> _members = 
            new List<QualifiedContext<ParserRuleContext>>();

        public ProcedureListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<ParserRuleContext>> Members { get { return _members; } }

        private void AddMember<TContext>(TContext context) where TContext : ParserRuleContext
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            AddMember(context);
        }
    }
}