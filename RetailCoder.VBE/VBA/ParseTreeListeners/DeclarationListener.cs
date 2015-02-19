using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class DeclarationListener : VBListenerBase, IExtensionListener<ParserRuleContext>
    {
        private readonly QualifiedModuleName _qualifiedName;

        private readonly IList<QualifiedContext<ParserRuleContext>> _members = 
            new List<QualifiedContext<ParserRuleContext>>();

        public DeclarationListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<ParserRuleContext>> Members { get { return _members; } }

        public override void EnterVariableStmt(VBParser.VariableStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
            foreach (var child in context.VariableListStmt().VariableSubStmt())
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, child));
            }
        }

        public override void EnterEnumerationStmt(VBParser.EnumerationStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterConstStmt(VBParser.ConstStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
            foreach (var child in context.ConstSubStmt())
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, child));
            }
        }

        public override void EnterTypeStmt(VBParser.TypeStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterDeclareStmt(VBParser.DeclareStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterEventStmt(VBParser.EventStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterArg(VBParser.ArgContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }
    }

    public class DeclarationSectionListener : DeclarationListener
    {
        public DeclarationSectionListener(QualifiedModuleName qualifiedName)
            : base(qualifiedName)
        {
        }

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
