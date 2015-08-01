using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.ParserErrors;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IRubberduckMenuFactory _rubberduckMenuFactory;
        private readonly IParserErrorsPresenterFactory _parserErrorsPresenterFactory;
        private readonly IRubberduckParserFactory _parserFactory;
        private readonly IInspectorFactory _inspectorFactory;
        private IInspector _inspector;
        private IParserErrorsPresenter _parserErrorsPresenter;
        private readonly IGeneralConfigService _configService;
        private readonly IActiveCodePaneEditor _editor;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private IRubberduckParser _parser;

        private Configuration _config;
        private IRubberduckMenu _menu;
        private FormContextMenu _formContextMenu;
        private CodeInspectionsToolbar _codeInspectionsToolbar;

        private bool _displayToolbar;
        private Point _toolbarCoords = new Point(-1, -1);

        public App(VBE vbe, 
            IMessageBox messageBox,
            IRubberduckMenuFactory rubberduckMenuFactory,
            IParserErrorsPresenterFactory parserErrorsPresenterFactory,
            IRubberduckParserFactory parserFactory,
            IInspectorFactory inspectorFactory, 
            IGeneralConfigService configService, 
            ICodePaneWrapperFactory wrapperFactory, 
            IActiveCodePaneEditor editor)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _rubberduckMenuFactory = rubberduckMenuFactory;
            _parserErrorsPresenterFactory = parserErrorsPresenterFactory;
            _parserFactory = parserFactory;
            _inspectorFactory = inspectorFactory;
            _configService = configService;

            _configService.SettingsChanged += _configService_SettingsChanged;

            _editor = editor;
            _wrapperFactory = wrapperFactory;
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

            _inspector = _inspectorFactory.Create();

            _parserErrorsPresenter = _parserErrorsPresenterFactory.Create();

            _menu = _rubberduckMenuFactory.Create();
            _menu.Initialize();

            _formContextMenu = new FormContextMenu(_vbe, _parser, _editor, _wrapperFactory);
            _formContextMenu.Initialize();

            _codeInspectionsToolbar = new CodeInspectionsToolbar(_vbe, _inspector);
            _codeInspectionsToolbar.Initialize();

            if (_toolbarCoords.X != -1 && _toolbarCoords.Y != -1)
            {
                _codeInspectionsToolbar.ToolbarCoords = _toolbarCoords;
            }
            _codeInspectionsToolbar.ToolbarVisible = _displayToolbar;
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

            if (_formContextMenu != null)
            {
                _formContextMenu.Dispose();
            }

            if (_codeInspectionsToolbar != null)
            {
                _displayToolbar = _codeInspectionsToolbar.ToolbarVisible;
                _toolbarCoords = _codeInspectionsToolbar.ToolbarCoords;
                _codeInspectionsToolbar.Dispose();
            }

            if (_parser != null)
            {
                _parser.ParseStarted -= _parser_ParseStarted;
                _parser.ParserError -= _parser_ParserError;
            }
        }
    }
}
