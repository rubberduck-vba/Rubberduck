using Rubberduck.Settings;
using System;

namespace Rubberduck.SettingsProvider
{
    public class ConfigurationServiceBase<T> : IConfigurationService<T>
        where T : class, new()
    {
        private readonly IPersistenceService<T> _persister;
        protected readonly IDefaultSettings<T> Defaults;

        private readonly object valueLock = new object();
        protected T CurrentValue;

        public ConfigurationServiceBase(IPersistenceService<T> persister, IDefaultSettings<T> defaultSettings)
        {
            _persister = persister;
            Defaults = defaultSettings;
        }

        protected void OnSettingsChanged()
        {
            var eventArgs = new ConfigurationChangedEventArgs(false, false, false, false);
            SettingsChanged?.Invoke(this, eventArgs);
        }

        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;

        protected T LoadCacheValue()
        {
            lock(valueLock)
            {
                if (CurrentValue == null)
                {
                    T defaults = ReadDefaults();
                    T newValue = _persister.Load() ?? defaults;
                    CurrentValue = newValue;
                }
                return CurrentValue;
            }
        }

        public virtual T Read()
        {
            return LoadCacheValue();
        }

        public virtual T ReadDefaults()
        {
            return Defaults.Default;
        }

        protected void PersistValue(T settings)
        {
            lock (valueLock)
            {
                // purge current value
                CurrentValue = null;
                _persister.Save(settings);
            }
        }

        public virtual void Save(T settings)
        {
            PersistValue(settings);
            OnSettingsChanged();
        }

        public virtual T Import(string path)
        {
            T loaded = _persister.Load(path);
            Save(loaded);
            return Read();
        }

        public virtual void Export(string path)
        {
            T current = Read();
            _persister.Save(current, path);
        }
    }
}
