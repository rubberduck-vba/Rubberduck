using System;
using System.Globalization;
using System.Windows.Forms;
using NLog;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.ParserErrors;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly IMessageBox _messageBox;
        private readonly IParserErrorsPresenterFactory _parserErrorsPresenterFactory;
        private readonly IRubberduckParserFactory _parserFactory;
        private readonly IInspectorFactory _inspectorFactory;
        private readonly IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;

        private IParserErrorsPresenter _parserErrorsPresenter;
        private readonly Logger _logger;
        private IRubberduckParser _parser;

        private Configuration _config;

        public App(IMessageBox messageBox,
            IParserErrorsPresenterFactory parserErrorsPresenterFactory,
            IRubberduckParserFactory parserFactory,
            IInspectorFactory inspectorFactory, 
            IGeneralConfigService configService,
            IAppMenu appMenus)
        {
            _messageBox = messageBox;
            _parserErrorsPresenterFactory = parserErrorsPresenterFactory;
            _parserFactory = parserFactory;
            _inspectorFactory = inspectorFactory;
            _configService = configService;
            _appMenus = appMenus;
            _logger = LogManager.GetCurrentClassLogger();

            _configService.SettingsChanged += _configService_SettingsChanged;
        }

        public void Startup()
        {
            CleanReloadConfig();

            _appMenus.Initialize();
            _appMenus.Localize();
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
            _logger.Debug("Loading configuration");
            _config = _configService.LoadConfiguration();

            var currentCulture = RubberduckUI.Culture;
            try
            {
                RubberduckUI.Culture = CultureInfo.GetCultureInfo(_config.UserSettings.LanguageSetting.Code);
                _appMenus.Localize();
            }
            catch (CultureNotFoundException exception)
            {
                _logger.Error(exception, "Error Setting Culture for RubberDuck");
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

            _inspectorFactory.Create();

            _parserErrorsPresenter = _parserErrorsPresenterFactory.Create();
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
            if (_parser != null)
            {
                _parser.ParseStarted -= _parser_ParseStarted;
                _parser.ParserError -= _parser_ParserError;
            }
        }
    }
}
