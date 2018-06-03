using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class VBALogicOperators
    {
        public static bool Eqv(bool lhs, bool rhs) => !(lhs ^ rhs) || (lhs && rhs);

        public static int Eqv(int lhs, int rhs) => ~(lhs ^ rhs) | (lhs & rhs);

        public static long Eqv(long lhs, long rhs) => ~(lhs ^ rhs) | (lhs & rhs);

        public static bool Imp(bool lhs, bool rhs) => rhs || (!lhs && !rhs);

        public static int Imp(int lhs, int rhs) => rhs | (~lhs & ~rhs);

        public static long Imp(long lhs, long rhs) => rhs | (~lhs & ~rhs);

        //public static bool GT<T>(T lhs, T rhs) where T : IComparable<T>
        //    => lhs.CompareTo(rhs) > 0;

        //public static bool GTE<T>(T lhs, T rhs) where T : IComparable<T>
        //    => lhs.CompareTo(rhs) >= 0;

        //public static bool LT<T>(T lhs, T rhs) where T : IComparable<T>
        //    => lhs.CompareTo(rhs) < 0;

        //public static bool LTE<T>(T lhs, T rhs) where T : IComparable<T>
        //    => lhs.CompareTo(rhs) <= 0;

        //public static bool EQ<T>(T lhs, T rhs) where T : IComparable<T>
        //    => lhs.CompareTo(rhs) == 0;

        //public static bool NEQ<T>(T lhs, T rhs) where T : IComparable<T>
        //    => lhs.CompareTo(rhs) != 0;

        public static bool Like(string input, string pattern)
        {
            if (pattern.Equals("*"))
            {
                return true;
            }
            var regExpression = new StringBuilder(@"^" + pattern);
            regExpression.Replace("#", "[0-9]");
            regExpression.Replace("?", "+");
            regExpression.Replace("[!", "[^");
            regExpression.Replace("[+]", "[?]");
            regExpression.Replace("[[0-9]]", "[#]");
            if (!regExpression.ToString().EndsWith("*"))
            {
                regExpression.Append("$");
            }
            var regex = new Regex(regExpression.ToString());
            var result = regex.IsMatch(input);
            return result;
        }
    }
}
