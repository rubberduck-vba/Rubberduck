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
        private readonly AddIn _addIn;
        private Inspector _inspector;
        private IParserErrorsPresenter _parserErrorsPresenter;
        private readonly IConfigurationLoader _configService = new ConfigurationLoader();
        private readonly IActiveCodePaneEditor _editor;
        private readonly IRubberduckCodePaneFactory _factory;
        private readonly IRubberduckParser _parser;

        private Configuration _config;
        private RubberduckMenu _menu;
        private FormContextMenu _formContextMenu;
        private CodeInspectionsToolbar _codeInspectionsToolbar;
        private bool _displayToolbar = false;
        private Point _toolbarCoords = new Point(-1, -1);

        public App(VBE vbe, AddIn addIn, IParserErrorsPresenter presenter, IRubberduckParser parser, IRubberduckCodePaneFactory factory, IActiveCodePaneEditor editor)
        {
            _vbe = vbe;
            _addIn = addIn;
            _factory = factory;
            _parser = parser;

            _parserErrorsPresenter = presenter;
            _configService.SettingsChanged += _configService_SettingsChanged;

            // todo: figure out why Ninject can't seem to resolve the VBE dependency to ActiveCodePaneEditor if it's in the VBEDitor assembly.
            // could it be that the VBE type in the two assemblies is actually different? 
            // aren't the two assemblies using the exact same Microsoft.Vbe.Interop assemby?
            
            _editor = editor; // */ new ActiveCodePaneEditor(vbe, _factory);

            LoadConfig();

            CleanUp();

            Setup();
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            LoadConfig();

            CleanUp();

            Setup();
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
                MessageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.LanguageSetting.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
        }

        private void Setup()
        {
            //_parser = new RubberduckParser(_factory);
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParserError += _parser_ParserError;

            _inspector = new Inspector(_parser, _configService);

            _parserErrorsPresenter = new ParserErrorsPresenter(_vbe, _addIn);

            _menu = new RubberduckMenu(_vbe, _addIn, _configService, _parser, _editor, _inspector, _factory);
            _menu.Initialize();

            _formContextMenu = new FormContextMenu(_vbe, _parser, _editor, _factory);
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
            if (_menu != null)
            {
                _menu.Dispose();
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

            if (_inspector != null)
            {
                _inspector.Dispose();
            }

            if (_parser != null)
            {
                _parser.ParseStarted -= _parser_ParseStarted;
                _parser.ParserError -= _parser_ParserError;
            }
        }
    }
}
