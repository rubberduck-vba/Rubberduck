using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;
using Rubberduck.Inspections;
using Rubberduck.UI;
using Rubberduck.Config;
using Rubberduck.UI.CodeInspections;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck
{
    [ComVisible(false)]
    public class App : IDisposable
    {
        private readonly RubberduckMenu _menu;
        private readonly CodeInspectionsToolbar _codeInspectionsToolbar;
        private readonly IList<IInspection> _inspections;

        public App(VBE vbe, AddIn addIn)
        {
            var config = ConfigurationLoader.LoadConfiguration();
            _inspections = ConfigurationLoader.GetImplementedCodeInspections();

            EnableCodeInspections(config);
            var parser = new VBParser();

            _menu = new RubberduckMenu(vbe, addIn, config, parser, _inspections);
            _codeInspectionsToolbar = new CodeInspectionsToolbar(vbe, parser, _inspections);
        }

        private void EnableCodeInspections(Configuration config)
        {
            foreach (var inspection in _inspections)
            {           
                foreach (var setting in config.UserSettings.CodeInspectinSettings.CodeInspections)
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
