using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{

    public class ScopedDeclaration<TContext>
        where TContext : ParserRuleContext
    {
        private readonly TContext _context;

        public ScopedDeclaration(TContext context)
        {
            _context = context;
        }

        public TContext Context { get { return _context; } }
    }

    public class DeclarationListener : VisualBasic6BaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
        public IEnumerable<ParserRuleContext> Members { get { return _members; } }



        public override void EnterVariableStmt(VisualBasic6Parser.VariableStmtContext context)
        {
            _members.Add(context);
            foreach (var child in context.variableListStmt().variableSubStmt())
            {
                _members.Add(child);
            }
        }

        public override void EnterEnumerationStmt(VisualBasic6Parser.EnumerationStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterConstStmt(VisualBasic6Parser.ConstStmtContext context)
        {
            _members.Add(context);
            foreach (var child in context.constSubStmt())
            {
                _members.Add(child);
            }
        }

        public override void EnterTypeStmt(VisualBasic6Parser.TypeStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterDeclareStmt(VisualBasic6Parser.DeclareStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterEventStmt(VisualBasic6Parser.EventStmtContext context)
        {
            _members.Add(context);
        }

        public override void EnterArg(VisualBasic6Parser.ArgContext context)
        {
            _members.Add(context);
        }
    }
}