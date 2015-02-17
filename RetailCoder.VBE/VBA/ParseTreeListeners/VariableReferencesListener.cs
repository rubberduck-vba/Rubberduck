using System.Collections.Generic;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableReferencesListener : IVBBaseListener,
        IExtensionListener<VBParser.AmbiguousIdentifierContext>
    {
        private readonly IList<VBParser.AmbiguousIdentifierContext> _members = new List<VBParser.AmbiguousIdentifierContext>();

        public IEnumerable<VBParser.AmbiguousIdentifierContext> Members { get { return _members; } }

        public override void EnterAmbiguousIdentifier(VBParser.AmbiguousIdentifierContext context)
        {
            // exclude declarations
            if (!(context.Parent is VBParser.VariableSubStmtContext) &&
                !(context.Parent is VBParser.ConstSubStmtContext))
                _members.Add(context);
        }
    }
}