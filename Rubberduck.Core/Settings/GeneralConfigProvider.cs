using System;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class GeneralConfigProvider : ConfigurationServiceBase<GeneralSettings>
    {
        private GeneralSettings current;
        private readonly GeneralSettings defaultSettings;

        public GeneralConfigProvider(IPersistanceService<GeneralSettings> persister)
            : base(persister)
        {
            defaultSettings = new DefaultSettings<GeneralSettings, Properties.Settings>().Default;
        }

        public override GeneralSettings Load()
        {
            var updated = persister.Load(defaultSettings) ?? defaultSettings;

            CheckForEventsToRaise(updated);
            current = updated;

            return current;
        }

        public override GeneralSettings LoadDefaults()
        {
            return defaultSettings;
        }

        public override void Save(GeneralSettings settings)
        {
            CheckForEventsToRaise(settings);
            OnSettingsChanged();
            persister.Save(settings);
        }

        private void CheckForEventsToRaise(GeneralSettings other)
        {
            if (current == null || !Equals(other.Language, current.Language))
            {
                OnLanguageChanged(EventArgs.Empty);
            }
            if (current == null || 
                other.IsAutoSaveEnabled != current.IsAutoSaveEnabled || 
                other.AutoSavePeriod != current.AutoSavePeriod)
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