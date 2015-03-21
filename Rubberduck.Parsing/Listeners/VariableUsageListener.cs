using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Listeners
{
    public class VariableUsageListener : VBABaseListener, IExtensionListener<VBAParser.AmbiguousIdentifierContext>
    {
        private readonly IList<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _members = 
            new List<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

        private readonly QualifiedModuleName _qualifiedName;
        private QualifiedMemberName _currentMember;

        public VariableUsageListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
            _currentMember = new QualifiedMemberName(_qualifiedName, "(declarations)");
        }

        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        protected void AddMember(VBAParser.AmbiguousIdentifierContext context)
        {
            _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_currentMember, context));
        }

        public override void EnterICS_S_VariableOrProcedureCall(VBAParser.ICS_S_VariableOrProcedureCallContext context)
        {
            _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_currentMember, context.ambiguousIdentifier()));
        }

        public override void EnterWithStmt(VBAParser.WithStmtContext context)
        {
            _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_currentMember, context.implicitCallStmt_InStmt().iCS_S_VariableOrProcedureCall().ambiguousIdentifier()));
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }
    }
}