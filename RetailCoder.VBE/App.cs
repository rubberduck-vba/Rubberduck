using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.CodeInspections;
using Rubberduck.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly RubberduckMenu _menu;
        private readonly CodeInspectionsToolbar _codeInspectionsToolbar;
        private readonly IList<IInspection> _inspections;
        private readonly IGeneralConfigService _configService;

        public App(VBE vbe, AddIn addIn)
        {
            _configService = new ConfigurationLoader();
            _inspections = _configService.GetImplementedCodeInspections();

            var config = _configService.LoadConfiguration();
            EnableCodeInspections(config);
            var parser = new RubberduckParser();
            var editor = new ActiveCodePaneEditor(vbe);

            var inspector = new Inspector(parser, _inspections);
            _menu = new RubberduckMenu(vbe, addIn, _configService, parser, editor, inspector);
            _codeInspectionsToolbar = new CodeInspectionsToolbar(vbe, inspector);
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
            if (disposing && _menu != null)
            {
                _menu.Dispose();
            }

            if (disposing && _codeInspectionsToolbar != null)
            {
                _codeInspectionsToolbar.Dispose();
            }
        }

        public void CreateExtUi()
        {
            _menu.Initialize();
            _codeInspectionsToolbar.Initialize();
        }
    }
}
