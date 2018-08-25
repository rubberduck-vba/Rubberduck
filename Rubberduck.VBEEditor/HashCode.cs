using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace Rubberduck.VBEditor
{
    public static class HashCode
    {
        [SuppressMessage("ReSharper", "ForCanBeConvertedToForeach")]
        [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
        [SuppressMessage("ReSharper", "RedundantCast")]
        public static int Compute(params object[] values)
        {
            unchecked
            {
                const int initial = (int)2166136261;
                const int multiplier = (int)16777619;
                var hash = initial;
                for (var i = 0; i < values.Length; i++)
                {
                    hash = hash * multiplier + values[i].GetHashCode();
                }
                return hash;
            }
        }

        public static int Compute(IEnumerable<int> values)
        {
            return values.Aggregate(5381, (hash, value) => ((hash << 5) + hash) ^ value);
        }
    }
}