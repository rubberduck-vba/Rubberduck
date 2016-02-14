using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class GeneralSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public GeneralSettingsViewModel(Configuration config)
        {
            Languages = new ObservableCollection<DisplayLanguageSetting>(
                new[] 
            {
                new DisplayLanguageSetting("en-US"),
                new DisplayLanguageSetting("fr-CA"),
                new DisplayLanguageSetting("de-DE"),
                new DisplayLanguageSetting("sv-SE"),
                new DisplayLanguageSetting("ja-JP")
            });

            SelectedLanguage = Languages.First(l => l.Code == config.UserSettings.GeneralSettings.Language.Code);
        }

        public ObservableCollection<DisplayLanguageSetting> Languages { get; set; } 

        private DisplayLanguageSetting _selectedLanguage;
        public DisplayLanguageSetting SelectedLanguage
        {
            get { return _selectedLanguage; }
            set
            {
                if (!Equals(_selectedLanguage, value))
                {
                    _selectedLanguage = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<Hotkey> _hotkeys;
        public ObservableCollection<Hotkey> Hotkeys
        {
            get { return _hotkeys; }
            set
            {
                if (_hotkeys != value)
                {
                    _hotkeys = value;
                    OnPropertyChanged();
                }
            }
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.GeneralSettings.Language = SelectedLanguage;
        }

        public void SetToDefaults(Configuration config)
        {
            SelectedLanguage = Languages.First(l => l.Code == config.UserSettings.GeneralSettings.Language.Code);
        }
    }
}