namespace Rubberduck.Parsing
{
    public struct QualifiedMemberName
    {
        public QualifiedMemberName(QualifiedModuleName qualifiedModuleName, string member)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _member = member;
        }

        private readonly QualifiedModuleName _qualifiedModuleName;
        public QualifiedModuleName QualifiedModuleName { get { return _qualifiedModuleName; } }

        private readonly string _member;
        public string Name { get { return _member; } }

        public override int GetHashCode()
        {
            return (_qualifiedModuleName.GetHashCode().ToString() + _member).GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var other = (QualifiedMemberName)obj;
            return _qualifiedModuleName.Equals(other.QualifiedModuleName) && _member == other.Name;
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