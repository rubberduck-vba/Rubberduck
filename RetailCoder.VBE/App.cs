using System;
using System.Globalization;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.ParserErrors;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly IMessageBox _messageBox;
        private readonly IParserErrorsPresenterFactory _parserErrorsPresenterFactory;
        private readonly IRubberduckParserFactory _parserFactory;
        private readonly IInspectorFactory _inspectorFactory;
        private IParserErrorsPresenter _parserErrorsPresenter;
        private readonly IGeneralConfigService _configService;
        private readonly IRubberduckMenuFactory _menuFactory;
        
        private IRubberduckParser _parser;
        private IMenu _menu;
        private Configuration _config;

        public App(IMessageBox messageBox,
            //IMenu integratedUserInterface,
            IRubberduckMenuFactory menuFactory,
            IParserErrorsPresenterFactory parserErrorsPresenterFactory,
            IRubberduckParserFactory parserFactory,
            IInspectorFactory inspectorFactory, 
            IGeneralConfigService configService)
        {
            _messageBox = messageBox;
            _menuFactory = menuFactory;
            _parserErrorsPresenterFactory = parserErrorsPresenterFactory;
            _parserFactory = parserFactory;
            _inspectorFactory = inspectorFactory;
            _configService = configService;

            _configService.SettingsChanged += _configService_SettingsChanged;
        }

        public void Startup()
        {
            CleanReloadConfig();
        }

        private void CleanReloadConfig()
        {
            LoadConfig();
            CleanUp();
            Setup();
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            CleanReloadConfig();
        }

        private void LoadConfig()
        {
            _config = _configService.LoadConfiguration();

            var currentCulture = RubberduckUI.Culture;
            try
            {
                RubberduckUI.Culture = CultureInfo.GetCultureInfo(_config.UserSettings.LanguageSetting.Code);
            }
            catch (CultureNotFoundException exception)
            {
                _messageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.LanguageSetting.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
        }

        private void Setup()
        {
            _parser = _parserFactory.Create();
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParserError += _parser_ParserError;
            _parserErrorsPresenter = _parserErrorsPresenterFactory.Create();

            _inspectorFactory.Create();

            _menu = _menuFactory.Create();
            _menu.Initialize();
        }

        private void _parser_ParseStarted(object sender, ParseStartedEventArgs e)
        {
            _parserErrorsPresenter.Clear();
        }

        private void _parser_ParserError(object sender, ParseErrorEventArgs e)
        {
            _parserErrorsPresenter.AddError(e);
            _parserErrorsPresenter.Show();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            CleanUp();
        }

        private void CleanUp()
        {
            var menu = _menu as IDisposable;
            if (menu != null)
            {
                menu.Dispose();
            }

            if (_parser != null)
            {
                _parser.ParseStarted -= _parser_ParseStarted;
                _parser.ParserError -= _parser_ParserError;
            }
        }
    }
}
