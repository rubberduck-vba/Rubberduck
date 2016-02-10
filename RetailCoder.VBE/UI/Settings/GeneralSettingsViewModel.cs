using System;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum Languages
    {
        EN,
        FR,
        DE,
        SV,
        JA
    }

    public class GeneralSettingsViewModel : ViewModelBase
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public GeneralSettingsViewModel(IGeneralConfigService configService)
        {
            _configService = configService;
            _config = configService.LoadConfiguration();

            SelectedLanguage = (Languages)Enum.Parse(typeof(Languages), _config.UserSettings.LanguageSetting.Code.Substring(0, 2).ToUpper());
        }

        private Languages _selectedLanguage;
        public Languages SelectedLanguage
        {
            get { return _selectedLanguage; }
            set
            {
                if (_selectedLanguage != value)
                {
                    _selectedLanguage = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}