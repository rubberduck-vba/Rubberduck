using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.Parsing;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ProcedureListener : VBABaseListener, IExtensionListener<ParserRuleContext>
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

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            AddMember(context);
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            AddMember(context);
        }
    }
}