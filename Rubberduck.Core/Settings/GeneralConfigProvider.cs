using System;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class GeneralConfigProvider : ConfigurationServiceBase<GeneralSettings>
    {

        public GeneralConfigProvider(IPersistenceService<GeneralSettings> persister)
            : base(persister, new DefaultSettings<GeneralSettings, Properties.Settings>())
        {
        }

        public override GeneralSettings Read()
        {
            var before = CurrentValue;
            var updated = LoadCacheValue();
            CheckForEventsToRaise(before, updated);
            return updated;
        }

        public override void Save(GeneralSettings settings)
        {
            var before = CurrentValue;
            PersistValue(settings);
            CheckForEventsToRaise(before, settings);
            OnSettingsChanged();
        }

        private void CheckForEventsToRaise(GeneralSettings before, GeneralSettings after)
        {
            if (before == null || !Equals(after.Language, before.Language))
            {
                OnLanguageChanged(EventArgs.Empty);
            }
            if (before == null ||
                after.IsAutoSaveEnabled != before.IsAutoSaveEnabled ||
                after.AutoSavePeriod != before.AutoSavePeriod)
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