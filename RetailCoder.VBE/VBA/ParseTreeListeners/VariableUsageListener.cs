using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableUsageListener : VBListenerBase, IExtensionListener<VBParser.AmbiguousIdentifierContext>
    {
        private readonly IList<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _members = 
            new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

        private readonly QualifiedModuleName _qualifiedName;
        private QualifiedMemberName _currentMember;

        public VariableUsageListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
            _currentMember = new QualifiedMemberName(_qualifiedName, "(declarations)");
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        protected void AddMember(VBParser.AmbiguousIdentifierContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_currentMember, context));
        }

        public override void EnterForNextStmt(VBParser.ForNextStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_currentMember, context.AmbiguousIdentifier().First()));
        }

        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_currentMember, context.AmbiguousIdentifier()));
        }

        public override void EnterWithStmt(VBParser.WithStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_currentMember, context.ImplicitCallStmt_InStmt().ICS_S_VariableCall().VariableCallStmt().AmbiguousIdentifier()));
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.AmbiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.AmbiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.AmbiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.AmbiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.AmbiguousIdentifier().GetText());
        }
    }
}