using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Rubberduck.Settings
{
    public class DefaultSettings<T, S>
        where S : System.Configuration.ApplicationSettingsBase
    {
        public IEnumerable<T> Defaults { get; }
        public T Default => Defaults.First();

        public DefaultSettings()
        {
            var properties = typeof(S).GetProperties().Where(prop => prop.PropertyType == typeof(T));
            var defaultInstance = typeof(S).GetProperty("Default", BindingFlags.Static | BindingFlags.Public).GetValue(null);

            Defaults = properties.Select(prop => prop.GetValue(defaultInstance)).Cast<T>();
        }
    }
}
