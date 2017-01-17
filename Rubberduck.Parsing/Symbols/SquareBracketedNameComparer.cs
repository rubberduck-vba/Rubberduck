using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public class SquareBracketedNameComparer : IEqualityComparer<string>
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
            if (value == null) return null;

            return value[0] == '[' && value[value.Length - 1] == ']'
                ? value
                : "[" + value + "]";
        }
    }
}