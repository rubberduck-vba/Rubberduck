////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

namespace Rubberduck.RibbonDispatcher.Concrete {
    internal static class RibbonCommonExtensions {
        public static TValue GetOrDefault<TValue>(this IReadOnlyDictionary<string, TValue> dictionary, string key) {
            TValue ctrl;
            return dictionary.TryGetValue(key??"", out ctrl) ? ctrl : default(TValue);
        }
    }
}
