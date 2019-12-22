using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.VBEditor.Utility
{
    public static class EnumHelper
    {
        public static Dictionary<TUnderlying, string> ToDictionary<TEnum, TUnderlying>()
        {
            var enumType = typeof(TEnum);
            var underlyingType = typeof(TUnderlying);
            var actualUnderlyingType = Enum.GetUnderlyingType(enumType);

            Debug.Assert(enumType.IsEnum, $"Type '{enumType.Name}' is not an enum");
            Debug.Assert(actualUnderlyingType == underlyingType, $"Underlying type parameter '{underlyingType.Name}' does not match actual underlying type for enum '{enumType.Name}' ('{actualUnderlyingType.Name}')");
            
            var dictionary = new Dictionary<TUnderlying, string>();

            foreach (var fieldInfo in enumType.GetFields().Where(fi => fi.FieldType.IsEnum))
            {
                if (fieldInfo.GetCustomAttributes(typeof(ReflectionIgnoreAttribute), false).Any())
                {
                    continue;
                }

                var key = (TUnderlying) fieldInfo.GetRawConstantValue();

                dictionary[key] = dictionary.ContainsKey(key)
                    ? $"{dictionary[key]} / {fieldInfo.Name}"
                    : fieldInfo.Name;
            }

            return dictionary;
        }
    } 

    public class ReflectionIgnoreAttribute : Attribute
    {
    }
}
