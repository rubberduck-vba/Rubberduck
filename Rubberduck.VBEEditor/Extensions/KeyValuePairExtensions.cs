using System.Collections.Generic;

namespace Rubberduck.VBEditor.Extensions
{
    public static class KeyValuePairExtensions
    {
        //See https://stackoverflow.com/a/43282724/5536802
        public static void Deconstruct<TKey, TValue>(this KeyValuePair<TKey, TValue> kvp, out TKey key, out TValue value)
        {
            key = kvp.Key;
            value = kvp.Value;
        }
    }
}
