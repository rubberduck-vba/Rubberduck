using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
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

    public class ScopedDeclaration<TContext>
        where TContext : ParserRuleContext
    {
        private readonly TContext _context;
        private readonly QualifiedMemberName _scope;

        public ScopedDeclaration(TContext context, QualifiedModuleName scope)
            :this(context, new QualifiedMemberName(scope, string.Empty))
        {
        }

        public ScopedDeclaration(TContext context, QualifiedMemberName scope)
        {
            _context = context;
            _scope = scope;
        }

        public TContext Context { get { return _context; } }
        public QualifiedMemberName Scope { get { return _scope; } }
    }
}