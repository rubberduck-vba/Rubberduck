namespace Rubberduck.Inspections
{
    public struct QualifiedMemberName
    {
        public QualifiedMemberName(QualifiedModuleName moduleScope, string member)
        {
            _moduleScope = moduleScope;
            _member = member;
        }

        private readonly QualifiedModuleName _moduleScope;
        public QualifiedModuleName ModuleScope { get { return _moduleScope; } }

        private readonly string _member;
        public string Name { get { return _member; } }

        public override int GetHashCode()
        {
            return (_moduleScope.GetHashCode().ToString() + _member).GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var other = (QualifiedMemberName)obj;
            return _moduleScope.Equals(other.ModuleScope) && _member == other.Name;
        }

        public static bool operator ==(QualifiedMemberName a, QualifiedMemberName b)
        {
            return a.Equals(b);
        }

        public static bool operator !=(QualifiedMemberName a, QualifiedMemberName b)
        {
            return !a.Equals(b);
        }
    }
}