using System.Collections.Generic;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableReferencesListener : VBListenerBase,
    IExtensionListener<VBParser.AmbiguousIdentifierContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _members = 
            new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

        public VariableReferencesListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        public override void EnterAmbiguousIdentifier(VBParser.AmbiguousIdentifierContext context)
        {
            // exclude declarations
            if (!(context.Parent is VBParser.VariableSubStmtContext) &&
                !(context.Parent is VBParser.ConstSubStmtContext))
                _members.Add(new QualifiedContext<VBParser.AmbiguousIdentifierContext>(_qualifiedName, context));
        }
    }
}