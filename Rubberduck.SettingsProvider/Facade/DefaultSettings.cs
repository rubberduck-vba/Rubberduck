using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Rubberduck.Settings
{
    public interface IDefaultSettings<T>
    {
        IEnumerable<T> Defaults { get; }
        T Default { get; }
    }

    public class FixedValueDefault<T> : IDefaultSettings<T>
    {
        public IEnumerable<T> Defaults { get => new[] { Default }; }
        public T Default { get; }

        public FixedValueDefault(T value)
        {
            Default = value;
        }
    }

    public class DefaultSettings<T, S>  : IDefaultSettings<T>
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

        public DefaultSettings(S settingsInstance)
        {
            var properties = typeof(S).GetProperties().Where(prop => prop.PropertyType == typeof(T));
            Defaults = properties.Select(prop => prop.GetValue(settingsInstance)).Cast<T>();
        }
    }
}
