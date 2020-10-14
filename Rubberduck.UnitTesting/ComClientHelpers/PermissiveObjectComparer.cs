using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.Variants;

namespace Rubberduck.UnitTesting.ComClientHelpers
{
    public class PermissiveObjectComparer : IEqualityComparer<object>
    {
        /// <summary>
        /// Tests equity between 2 objects using VBA's type promotion rules.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns>VBA equity</returns>
        public new bool Equals(object x, object y)
        {
            return VariantComparer.Compare(x, y) == VariantComparisonResults.VARCMP_EQ;
        }

        /// <summary>
        /// DO NOT USE THIS.  It is a hard-coded throw.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns>NotSupportedException</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1065")]     // see comment below.
        public int GetHashCode(object obj)
        {
            //This is intentional to "discourage" any use of the comparer that relies on GetHashCode().
            throw new NotSupportedException();
        }
    }
}
