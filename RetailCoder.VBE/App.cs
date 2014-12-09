using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;
using Rubberduck.Inspections;
using Rubberduck.UI;
using Rubberduck.Config;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;
using Rubberduck.Extensions;

namespace Rubberduck
{
    [ComVisible(false)]
    public class App : IDisposable
    {
        private readonly RubberduckMenu _menu;
        private readonly IList<IInspection> _inspections;

        public App(VBE vbe, AddIn addIn)
        {
            HostApplication.Type = vbe.HostApplication();

            var config = ConfigurationLoader.LoadConfiguration();

            var grammar = Assembly.GetExecutingAssembly()
                                  .GetTypes()
                                  .Where(type => type.BaseType == typeof(SyntaxBase))
                                  .Select(type =>
                                  {
                                      var constructorInfo = type.GetConstructor(Type.EmptyTypes);
                                      return constructorInfo != null ? constructorInfo.Invoke(Type.EmptyTypes) : null;
                                  })
                                  .Where(syntax => syntax != null)
                                  .Cast<ISyntax>()
                                  .ToList();

            _inspections = Assembly.GetExecutingAssembly()
                                   .GetTypes()
                                   .Where(type => type.GetInterfaces().Contains(typeof(IInspection)))
                                   .Select(type =>
                                   {
                                       var constructor = type.GetConstructor(Type.EmptyTypes);
                                       return constructor != null ? constructor.Invoke(Type.EmptyTypes) : null;
                                   })
                                  .Where(inspection => inspection != null)
                                   .Cast<IInspection>()
                                   .ToList();

            EnableCodeInspections();
            var parser = new Parser(grammar);

            _menu = new RubberduckMenu(vbe, addIn, config, parser, _inspections);
        }

        private void EnableCodeInspections()
        {
            foreach (var inspection in _inspections)
            {
                // todo: fetch value from configuration
                // inspection.Severity = ??;
            }
        }

        public void Dispose()
        {
            _menu.Dispose();
        }

        public void CreateExtUi()
        {
            _menu.Initialize();
        }
    }
}
