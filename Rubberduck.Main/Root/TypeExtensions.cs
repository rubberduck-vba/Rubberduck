using Rubberduck.Parsing.Common;
using Rubberduck.Settings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Root
{
    internal static class TypeExtensions
    {
        internal static bool NotDisabledExperimental(this Type type, GeneralSettings initialSettings)
        {
            var attribute = type.CustomAttributes.FirstOrDefault(f => f.AttributeType == typeof(ExperimentalAttribute));
            var ctorArg = attribute?.ConstructorArguments.Any() == true ? (string)attribute.ConstructorArguments.First().Value : string.Empty;

            return attribute == null || initialSettings.EnableExperimentalFeatures.Any(a => a.Key == ctorArg && a.IsEnabled);
        }

        internal static bool IsBasedOn(this Type type, Type allegedBase)
        {
            return allegedBase.IsAssignableFrom(type);
        }
    }
}
