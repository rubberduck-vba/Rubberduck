using Antlr4.Runtime;
using Rubberduck.Inspections;

namespace Rubberduck.VBA.ParseTreeListeners
{
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