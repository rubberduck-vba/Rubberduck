using System;
using System.Globalization;
using System.Windows.Threading;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.UI.Settings
{
    public class Settings : IDisposable
    {
        private static IConfigurationService<Configuration> _configService;
        private static CultureInfo _cultureInfo = null;

        public Settings(IConfigurationService<Configuration> configService)
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
                _cultureInfo = Resources.RubberduckUI.Culture ?? Dispatcher.CurrentDispatcher.Thread.CurrentUICulture;
                return;
            }

            try
            {
                var config = _configService.Read();
                _cultureInfo = CultureInfo.GetCultureInfo(config.UserSettings.GeneralSettings.Language.Code);
                
                Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = _cultureInfo;
            }
            catch (CultureNotFoundException)
            {
                _cultureInfo = Resources.RubberduckUI.Culture;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= SettingsChanged;
            }
            _isDisposed = true;
        }
    }
}