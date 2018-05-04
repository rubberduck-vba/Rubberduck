using Rubberduck.Parsing.Common;
using Rubberduck.Settings;
using System;
using System.Linq;

namespace Rubberduck.Root
{
    internal static class TypeExtensions
    {
        internal static bool NotDisabledOrExperimental(this Type type, GeneralSettings initialSettings)
        {
            return type.NotDisabled() && type.NotExperimental(initialSettings);
        }

        internal static bool NotExperimental(this Type type, GeneralSettings initialSettings)
        {
            var attribute = type.CustomAttributes.FirstOrDefault(f => f.AttributeType == typeof(ExperimentalAttribute));
            var ctorArg = attribute?.ConstructorArguments.Any() == true ? (string)attribute.ConstructorArguments.First().Value : string.Empty;

            return attribute == null || initialSettings.EnableExperimentalFeatures.Any(a => a.Key == ctorArg && a.IsEnabled);
        }

        internal static bool NotDisabled(this Type type)
        {
            return !Attribute.IsDefined(type, typeof(DisabledAttribute));
        }

        internal static bool IsBasedOn(this Type type, Type allegedBase)
        {
            return allegedBase.IsAssignableFrom(type);
        }
    }
}
