﻿using System.Collections.Generic;

namespace Rubberduck.RibbonDispatcher.Abstract {
    internal static class RibbonCommonExtensions {
        public static TValue GetOrDefault<TValue>(this IReadOnlyDictionary<string, TValue> dictionary, string key)
            where TValue : IRibbonCommon {
            TValue ctrl;
            return dictionary.TryGetValue(key, out ctrl) ? ctrl : default(TValue);
        }
    }
}
