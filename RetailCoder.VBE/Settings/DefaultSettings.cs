using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Settings
{
    public class DefaultSettings<T>
    {
        public IEnumerable<T> Defaults { get; }

        public DefaultSettings()
        {
            var properties = typeof(Properties.Settings).GetProperties().Where(prop => prop.PropertyType == typeof(T));

            Defaults = properties.Select(prop => prop.GetValue(Properties.Settings.Default)).Cast<T>();
        }
    }
}
