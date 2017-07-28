using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class QualifiedContext
    {
        public QualifiedContext(QualifiedMemberName memberName, ParserRuleContext context)
        {
            _memberName = memberName;
            _moduleName = memberName.QualifiedModuleName;
            _context = context;
        }

        public QualifiedContext(QualifiedModuleName moduleName, ParserRuleContext context)
            : this(moduleName.QualifyMemberName(moduleName.Name), context)
        {
        }

        private readonly QualifiedMemberName _memberName;
        public QualifiedMemberName MemberName { get { return _memberName; } }

        private readonly QualifiedModuleName _moduleName;
        public QualifiedModuleName ModuleName { get { return _moduleName; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        public override int GetHashCode()
        {
            return Context.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var other = obj as QualifiedContext;
            if (other == null)
            {
                return false;
            }

            return other.ModuleName == ModuleName && other.MemberName == MemberName;
        }

        public static bool operator ==(QualifiedContext context1, QualifiedContext context2)
        {
            if (((object)context1) == null)
            {
                return ((object)context2) == null;
            }

            return context1.Equals(context2);
        }

        public static bool operator !=(QualifiedContext context1, QualifiedContext context2)
        {
            if (((object)context1) == null)
            {
                return ((object)context2) != null;
            }

            return !context1.Equals(context2);
        }
    }

    public class QualifiedContext<TContext> : QualifiedContext
        where TContext : ParserRuleContext
    {
        public QualifiedContext(QualifiedMemberName memberName, TContext context)
            : base(memberName, context)
        {
        }

        public QualifiedContext(QualifiedModuleName qualifiedName, TContext context)
            :base(qualifiedName, context)
        {
        }

        public new TContext Context => base.Context as TContext;
    }
}
