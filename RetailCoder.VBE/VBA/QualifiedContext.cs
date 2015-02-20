using System.Runtime.InteropServices;
using Rubberduck.Inspections;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public class QualifiedContext<TContext>
    {
        public QualifiedContext(QualifiedMemberName memberName, TContext context)
            : this(memberName.ModuleScope, context)
        {
            _member = memberName;
        }

        public QualifiedContext(QualifiedModuleName qualifiedName, TContext context)
        {
            _qualifiedName = qualifiedName;
            _context = context;
        }

        private readonly QualifiedMemberName _member;
        public QualifiedMemberName MemberName { get { return _member; } }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly TContext _context;
        public TContext Context { get { return _context; } }
    }
}
