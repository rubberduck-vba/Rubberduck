using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class SquareBracketedNameComparer : 
        IEqualityComparer<string>, 
        IEqualityComparer<QualifiedMemberName>, 
        IEqualityComparer<QualifiedModuleName>,
        IEqualityComparer<Declaration>
    {
        public bool Equals(string x, string y)
        {
            return string.Equals(ApplyBrackets(x), ApplyBrackets(y));
        }

        public int GetHashCode(string obj)
        {
            if (obj == null) return 0;

            var value = ApplyBrackets(obj);
            return value.GetHashCode();
        }

        private string ApplyBrackets(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            return value[0] == '[' && value[value.Length - 1] == ']'
                ? value
                : "[" + value + "]";
        }

        public bool Equals(QualifiedMemberName x, QualifiedMemberName y)
        {
            return Equals(x.MemberName, y.MemberName);
        }

        public int GetHashCode(QualifiedMemberName obj)
        {
            return GetHashCode(obj.MemberName);
        }

        public bool Equals(QualifiedModuleName x, QualifiedModuleName y)
        {
            return Equals(x.Name, y.Name);
        }

        public int GetHashCode(QualifiedModuleName obj)
        {
            return GetHashCode(obj.Name);
        }

        public bool Equals(Declaration x, Declaration y)
        {
            return Equals(x.IdentifierName, y.IdentifierName);
        }

        public int GetHashCode(Declaration obj)
        {
            return GetHashCode(obj.IdentifierName);
        }
    }
}