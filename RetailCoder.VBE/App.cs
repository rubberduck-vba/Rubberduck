using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.ParserErrors;
using Rubberduck.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly IRubberduckParser _parser;
        private readonly RubberduckMenu _menu;
        private readonly CodeInspectionsToolbar _codeInspectionsToolbar;
        private readonly IList<IInspection> _inspections;
        private readonly IGeneralConfigService _configService;
        private readonly Inspector _inspector;
        private readonly ParserErrorsPresenter _parserErrorsPresenter;

        public App(VBE vbe, AddIn addIn)
        {
            _configService = new ConfigurationLoader();
            _inspections = _configService.GetImplementedCodeInspections();

            var config = _configService.LoadConfiguration();
            RubberduckUI.Culture = CultureInfo.GetCultureInfo(config.UserSettings.LanguageSetting.Code);

            EnableCodeInspections(config);
            _parser = new RubberduckParser();
            _parserErrorsPresenter = new ParserErrorsPresenter(vbe, addIn);
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParserError += _parser_ParserError;

            var editor = new ActiveCodePaneEditor(vbe);

            _inspector = new Inspector(_parser, _inspections);
            _menu = new RubberduckMenu(vbe, addIn, _configService, _parser, editor, _inspector);
            _codeInspectionsToolbar = new CodeInspectionsToolbar(vbe, _inspector);
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
                    if (inspection.Name == setting.Name)
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

        public void CreateExtUi()
        {
            _menu.Initialize();
            _codeInspectionsToolbar.Initialize();
        }
    }
}
