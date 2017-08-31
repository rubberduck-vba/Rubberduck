using System.Collections.Generic;

namespace Rubberduck.RibbonDispatcher {
    /// <summary>TODO</summary>
    public static class DictionaryExtensions {
        /// <summary>Adds the specified element to the dictionary only when it is not null.</summary>
        public static void AddNotNull<TValue>(this IDictionary<string, TValue> dictionary, string itemId, TValue ctrl) {
            if (ctrl != null) { dictionary?.Add(itemId, ctrl); }
        }

        /// <summary>TODO</summary>
        public static TValue GetOrDefault<TValue>(this IReadOnlyDictionary<string, TValue> dictionary, string key) {
            if (dictionary == null) return default(TValue);
            TValue ctrl;
            return dictionary.TryGetValue(key??"", out ctrl) ? ctrl : default(TValue);
        }
    }
}
