using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Listeners
{
    public class DeclarationListener : VBABaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly QualifiedModuleName _qualifiedName;

        private readonly IList<QualifiedContext<ParserRuleContext>> _members = 
            new List<QualifiedContext<ParserRuleContext>>();

        public DeclarationListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<ParserRuleContext>> Members { get { return _members; } }

        public override void EnterVariableStmt(VBAParser.VariableStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
            foreach (var child in context.variableListStmt().variableSubStmt())
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, child));
            }
        }

        public override void EnterVisibility(VBAParser.VisibilityContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterConstStmt(VBAParser.ConstStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
            foreach (var child in context.constSubStmt())
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, child));
            }
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterArg(VBAParser.ArgContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }
    }
}
