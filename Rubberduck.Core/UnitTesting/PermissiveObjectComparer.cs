using System;
using System.Collections.Generic;

namespace Rubberduck.UnitTesting
{
    internal class PermissiveObjectComparer : IEqualityComparer<object>
    {
        /// <summary>
        /// Tests equity between 2 objects using VBA's type promotion rules.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns>VBA equity</returns>
        public new bool Equals(object x, object y)
        {
            var expected = x;
            var actual = y;

            // try promoting integral types first.
            if (expected is ulong && actual is ulong)
            {
                return (ulong)x == (ulong)y;
            }
            // then try promoting to floating point
            if (expected is double && actual is double)
            {
                // ReSharper disable once CompareOfFloatsByEqualityOperator - We're cool with that.
                return (double)x == (double)y;
            }
            // that shouldn't actually happen, since decimal is the only numeric ValueType in its category
            // this means we should've gotten the same types earlier in the Assert method
            if (expected is decimal && actual is decimal)
            {
                return (decimal)x == (decimal)y;
            }
            // worst case scenario for numbers
            // since we're inside VBA though, double is the more appropriate type to compare, 
            // because that is what's used internally anyways, see https://support.microsoft.com/en-us/kb/78113
            if ((expected is decimal && actual is double) || (expected is double && actual is decimal))
            {
                // ReSharper disable once CompareOfFloatsByEqualityOperator - We're still cool with that.
                return (double)x == (double)y;
            }
            // no number-type promotions are applicable. 2nd to last straw: string "promotion"
            if (expected is string || actual is string)
            {
                expected = expected.ToString();
                actual = actual.ToString();
                return expected.Equals(actual);
            }
            return x.Equals(y);
        }

        /// <summary>
        /// DO NOT USE THIS.  It is a hard-coded throw.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns>NotSupportedException</returns>
        public int GetHashCode(object obj)
        {
            //This is intentional to "discourage" any use of the comparer that relies on GetHashCode().
            throw new NotSupportedException();
        }
    }
}
