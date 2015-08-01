using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.ParserErrors;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly IMessageBox _messageBox;
        private readonly IRubberduckMenuFactory _rubberduckMenuFactory;
        private readonly IParserErrorsPresenterFactory _parserErrorsPresenterFactory;
        private readonly IRubberduckParserFactory _parserFactory;
        private readonly IInspectorFactory _inspectorFactory;
        private IParserErrorsPresenter _parserErrorsPresenter;
        private readonly IGeneralConfigService _configService;
        private IRubberduckParser _parser;

        private Configuration _config;
        private IMenu _menu;
        private readonly IMenu _formContextMenu;
        private readonly IToolbar _codeInspectionsToolbar;

        private bool _displayToolbar;
        private Point _toolbarLocation = new Point(-1, -1);

        public App(IMessageBox messageBox,
            IRubberduckMenuFactory rubberduckMenuFactory,
            IParserErrorsPresenterFactory parserErrorsPresenterFactory,
            IRubberduckParserFactory parserFactory,
            IInspectorFactory inspectorFactory, 
            IGeneralConfigService configService,
            [FormContextMenu] IMenu formContextMenu,
            [CodeInspectionsToolbar] IToolbar codeInspectionsToolbar)
        {
            _messageBox = messageBox;
            _rubberduckMenuFactory = rubberduckMenuFactory;
            _parserErrorsPresenterFactory = parserErrorsPresenterFactory;
            _parserFactory = parserFactory;
            _inspectorFactory = inspectorFactory;
            _configService = configService;

            _configService.SettingsChanged += _configService_SettingsChanged;

            _formContextMenu = formContextMenu;
            _codeInspectionsToolbar = codeInspectionsToolbar;
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

            _inspectorFactory.Create();

            _parserErrorsPresenter = _parserErrorsPresenterFactory.Create();

            _menu = _rubberduckMenuFactory.Create();
            _menu.Initialize();

            _formContextMenu.Initialize();
            _codeInspectionsToolbar.Initialize();

            if (_toolbarLocation.X != -1 && _toolbarLocation.Y != -1)
            {
                _codeInspectionsToolbar.Location = _toolbarLocation;
            }
            _codeInspectionsToolbar.Visible = _displayToolbar;
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

            var formContextMenu = _formContextMenu as IDisposable;
            if (formContextMenu != null)
            {
                formContextMenu.Dispose();
            }

            _displayToolbar = _codeInspectionsToolbar.Visible;
            _toolbarLocation = _codeInspectionsToolbar.Location;

            var codeInspectionsToolbar = _codeInspectionsToolbar as IDisposable;
            if (codeInspectionsToolbar != null)
            {
                codeInspectionsToolbar.Dispose();
            }

            if (_parser != null)
            {
                _parser.ParseStarted -= _parser_ParseStarted;
                _parser.ParserError -= _parser_ParserError;
            }
        }
    }
}
