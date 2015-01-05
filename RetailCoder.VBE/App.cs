using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;
using Rubberduck.Inspections;
using Rubberduck.UI;
using Rubberduck.Config;
using Rubberduck.UI.CodeInspections;
using Rubberduck.VBA.Parser;

namespace Rubberduck
{
    [ComVisible(false)]
    public class App : IDisposable
    {
        private readonly RubberduckMenu _menu;
        private readonly CodeInspectionsToolbar _codeInspectionsToolbar;
        private readonly IList<IInspection> _inspections;
        private readonly IConfigurationService _configService;

        public App(VBE vbe, AddIn addIn)
        {
            _configService = new ConfigurationLoader();

            var grammar = _configService.GetImplementedSyntax();

            _inspections = _configService.GetImplementedCodeInspections();

            var config = _configService.LoadConfiguration();
            EnableCodeInspections(config);

            var parser = new Parser(grammar);

            _menu = new RubberduckMenu(vbe, addIn, _configService, parser, _inspections);
            _codeInspectionsToolbar = new CodeInspectionsToolbar(vbe, parser, _inspections);
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
            _menu.Dispose();
        }

        public void CreateExtUi()
        {
            _menu.Initialize();
            _codeInspectionsToolbar.Initialize();
        }
    }
}
