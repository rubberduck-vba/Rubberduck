using System.Collections.Generic;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class VariableReferencesListener : VisualBasic6BaseListener,
        IExtensionListener<VisualBasic6Parser.AmbiguousIdentifierContext>
    {
        private readonly IList<VisualBasic6Parser.AmbiguousIdentifierContext> _members = new List<VisualBasic6Parser.AmbiguousIdentifierContext>();

        public IEnumerable<VisualBasic6Parser.AmbiguousIdentifierContext> Members { get { return _members; } }

        public override void EnterAmbiguousIdentifier(VisualBasic6Parser.AmbiguousIdentifierContext context)
        {
            // exclude declarations
            if (!(context.Parent is VisualBasic6Parser.VariableSubStmtContext) &&
                !(context.Parent is VisualBasic6Parser.ConstSubStmtContext))
                _members.Add(context);
        }
    }
}