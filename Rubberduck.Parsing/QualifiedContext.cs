using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class QualifiedContext
    {
        public QualifiedContext(QualifiedMemberName memberName, ParserRuleContext context)
            : this(memberName.QualifiedModuleName, context)
        {
            MemberName = memberName;
        }

        public QualifiedContext(QualifiedModuleName moduleName, ParserRuleContext context)
        {
            ModuleName = moduleName;
            Context = context;
        }

        public QualifiedMemberName MemberName { get; }
        public QualifiedModuleName ModuleName { get; }
        public ParserRuleContext Context { get; }

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
            if (context1 is null)
            {
                return context2 is null;
            }

            return context1.Equals(context2);
        }

        public static bool operator !=(QualifiedContext context1, QualifiedContext context2)
        {
            if (context1 is null)
            {
                return !(context2 is null);
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
