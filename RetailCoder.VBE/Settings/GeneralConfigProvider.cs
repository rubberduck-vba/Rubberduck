using System;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class GeneralConfigProvider : IConfigProvider<GeneralSettings>
    {
        private GeneralSettings _current;
        private readonly IPersistanceService<GeneralSettings> _persister;
        private readonly GeneralSettings _defaultSettings;

        public GeneralConfigProvider(IPersistanceService<GeneralSettings> persister)
        {
            _persister = persister;
            _defaultSettings = new DefaultSettings<GeneralSettings>().Default;
        }

        public GeneralSettings Create()
        {
            var updated = _persister.Load(_defaultSettings) ?? _defaultSettings;

            CheckForEventsToRaise(updated);
            _current = updated;

            return _current;
        }

        public GeneralSettings CreateDefaults()
        {
            return _defaultSettings;
        }

        public void Save(GeneralSettings settings)
        {
            CheckForEventsToRaise(settings);
            _persister.Save(settings);
        }

        private void CheckForEventsToRaise(GeneralSettings other)
        {
            if (_current == null || !Equals(other.Language, _current.Language))
            {
                OnLanguageChanged(EventArgs.Empty);
            }
            if (_current == null || 
                other.IsAutoSaveEnabled != _current.IsAutoSaveEnabled || 
                other.AutoSavePeriod != _current.AutoSavePeriod)
            {
                OnAutoSaveSettingsChanged(EventArgs.Empty);
            }
        }

        public event EventHandler LanguageChanged;
        protected virtual void OnLanguageChanged(EventArgs e)
        {
            LanguageChanged?.Invoke(this, e);
        }

        public event EventHandler AutoSaveSettingsChanged;
        protected virtual void OnAutoSaveSettingsChanged(EventArgs e)
        {
            AutoSaveSettingsChanged?.Invoke(this, e);
        }
    }
}