using System;

namespace Rubberduck.VBEditor
{
    public struct QualifiedMemberName
    {
        public QualifiedMemberName(QualifiedModuleName qualifiedModuleName, string memberName)
        {
            _qualifiedModuleName = qualifiedModuleName;
            _memberName = memberName;
        }

        private readonly QualifiedModuleName _qualifiedModuleName;
        public QualifiedModuleName QualifiedModuleName { get { return _qualifiedModuleName; } }

        private readonly string _memberName;
        public string MemberName { get { return _memberName; } }

        public override string ToString()
        {
            return _qualifiedModuleName + "." + _memberName;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash * 23 + _qualifiedModuleName.GetHashCode();
                hash = hash * 23 + _memberName.GetHashCode();
                return hash;
            }
        }

        public override bool Equals(object obj)
        {
            try
            {
                var other = (QualifiedMemberName)obj;
                return _qualifiedModuleName.Equals(other.QualifiedModuleName) && _memberName == other.MemberName;
            }
            catch (InvalidCastException)
            {
                return false;
            }
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
