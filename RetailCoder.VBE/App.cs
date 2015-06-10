using System;
using System.Collections.Generic;
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

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private readonly IList<IInspection> _inspections;
        private readonly Inspector _inspector;
        private readonly ParserErrorsPresenter _parserErrorsPresenter;
        private readonly IGeneralConfigService _configService = new ConfigurationLoader();
        private readonly ActiveCodePaneEditor _editor;
        private readonly IRubberduckParser _parser;

        private Configuration _config;
        private RubberduckMenu _menu;
        private FormContextMenu _formContextMenu;
        private CodeInspectionsToolbar _codeInspectionsToolbar;

        public App(VBE vbe, AddIn addIn)
        {
            _vbe = vbe;
            _addIn = addIn;
            _inspections = _configService.GetImplementedCodeInspections();

            _parser = new RubberduckParser();
            _parserErrorsPresenter = new ParserErrorsPresenter(vbe, addIn);
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParserError += _parser_ParserError;
            _configService.SettingsChanged += _configService_SettingsChanged;

            _editor = new ActiveCodePaneEditor(vbe);

            _inspector = new Inspector(_parser, _inspections);

            LoadConfig();
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            LoadConfig();
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
                _configService.SaveConfiguration(_config, true);
            }

            EnableCodeInspections(_config);

            if (_menu != null)
            {
                _menu.Dispose();
            }
            
            if (_formContextMenu != null)
            {
                _formContextMenu.Dispose();
            }

            var displayToolbar = false;
            var toolbarCoords = new Point(-1, -1);
            if (_codeInspectionsToolbar != null)
            {
                displayToolbar = _codeInspectionsToolbar.ToolbarVisible;
                toolbarCoords = _codeInspectionsToolbar.ToolbarCoords;
                _codeInspectionsToolbar.Dispose();
            }

            _menu = new RubberduckMenu(_vbe, _addIn, _configService, _parser, _editor, _inspector);
            _menu.Initialize();

            _formContextMenu = new FormContextMenu(_vbe, _parser);
            _formContextMenu.Initialize();

            _codeInspectionsToolbar = new CodeInspectionsToolbar(_vbe, _inspector);
            _codeInspectionsToolbar.Initialize();

            if (toolbarCoords.X != -1 && toolbarCoords.Y != -1)
            {
                _codeInspectionsToolbar.ToolbarCoords = toolbarCoords;
            }
            _codeInspectionsToolbar.ToolbarVisible = displayToolbar;
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

        private void EnableCodeInspections(Configuration config)
        {
            foreach (var inspection in _inspections)
            {           
                foreach (var setting in config.UserSettings.CodeInspectionSettings.CodeInspections)
                {
                    if (inspection.Description == setting.Description)
                    {
                        inspection.Severity = setting.Severity;
                    }
                } 
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

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
                _codeInspectionsToolbar.Dispose();
            }

            if (_inspector != null)
            {
                _inspector.Dispose();
            }

            if (_parserErrorsPresenter != null)
            {
                _parserErrorsPresenter.Dispose();
            }
        }
    }
}
