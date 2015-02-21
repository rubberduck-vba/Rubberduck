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

        public VariableUsageListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        public override void EnterForNextStmt(VBParser.ForNextStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context.AmbiguousIdentifier().First()));
        }

        public override void EnterVariableCallStmt(VBParser.VariableCallStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context.AmbiguousIdentifier()));
        }

        public override void EnterWithStmt(VBParser.WithStmtContext context)
        {
            _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context.ImplicitCallStmt_InStmt().ICS_S_VariableCall().VariableCallStmt().AmbiguousIdentifier()));
        }
    }
}