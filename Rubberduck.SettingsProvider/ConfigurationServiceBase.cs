using System;

namespace Rubberduck.SettingsProvider
{
    public class ConfigurationServiceBase<T> : IConfigurationService<T>
        where T : class, new()
    {
        protected readonly IPersistanceService<T> persister;

        public ConfigurationServiceBase(IPersistanceService<T> persister)
        {
            this.persister = persister;
        }

        protected void OnSettingsChanged()
        {
            var eventArgs = new ConfigurationChangedEventArgs(false, false, false, false);
            SettingsChanged?.Invoke(this, eventArgs);
        }

        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;

        public virtual T Load()
        {
            var defaults = (T) typeof(T).GetConstructor(new Type[] { }).Invoke(new object[] { });
            return persister.Load(defaults) ?? defaults;
        }

        public virtual T LoadDefaults()
        {
            return (T)typeof(T).GetConstructor(new Type[] { }).Invoke(new object[] { });
        }

        public virtual void Save(T settings)
        {
            persister.Save(settings);
            OnSettingsChanged();
        }
    }
}
