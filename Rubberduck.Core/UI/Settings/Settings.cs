using System;
using System.Globalization;
using System.Windows.Threading;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public class Settings : IDisposable
    {
        private static IGeneralConfigService _configService;
        private static CultureInfo _cultureInfo = null;

        public Settings(IGeneralConfigService configService)
        {
            _configService = configService;
            _configService.SettingsChanged += SettingsChanged;
        }

        private void SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            if (e.LanguageChanged)
            {
                LoadLanguage();
            }
        }

        public static CultureInfo Culture
        {
            get
            {
                if (_cultureInfo != null)
                {
                    return _cultureInfo;
                }

                LoadLanguage();
                return _cultureInfo;
            }
        }

        private static void LoadLanguage()
        {
            if (_configService == null)
            {
                _cultureInfo = RubberduckUI.Culture ?? Dispatcher.CurrentDispatcher.Thread.CurrentUICulture;
                return;
            }

            try
            {
                var config = _configService.LoadConfiguration();
                _cultureInfo = CultureInfo.GetCultureInfo(config.UserSettings.GeneralSettings.Language.Code);
                
                Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = _cultureInfo;
            }
            catch (CultureNotFoundException)
            {
                _cultureInfo = RubberduckUI.Culture;
            }
        }

        public void Dispose()
        {
            if (_configService != null)
            {
                _configService.SettingsChanged -= SettingsChanged;
            }
        }
    }
}