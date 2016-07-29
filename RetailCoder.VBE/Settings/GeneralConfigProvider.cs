using System;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class GeneralConfigProvider : IConfigProvider<GeneralSettings>
    {
        private GeneralSettings _current;
        private readonly IPersistanceService<GeneralSettings> _persister;

        public GeneralConfigProvider(IPersistanceService<GeneralSettings> persister)
        {
            _persister = persister;
        }

        public GeneralSettings Create()
        {
            var prototype = new GeneralSettings();
            var updated = _persister.Load(prototype) ?? prototype;

            CheckForEventsToRaise(updated);
            _current = updated;

            return _current;
        }

        public GeneralSettings CreateDefaults()
        {
            return new GeneralSettings();
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
                other.AutoSaveEnabled != _current.AutoSaveEnabled || 
                other.AutoSavePeriod != _current.AutoSavePeriod)
            {
                OnAutoSaveSettingsChanged(EventArgs.Empty);
            }
        }

        public event EventHandler LanguageChanged;
        protected virtual void OnLanguageChanged(EventArgs e)
        {
            var handler = LanguageChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler AutoSaveSettingsChanged;
        protected virtual void OnAutoSaveSettingsChanged(EventArgs e)
        {
            var handler = AutoSaveSettingsChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}