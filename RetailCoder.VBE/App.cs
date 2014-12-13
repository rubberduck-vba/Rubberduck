using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;
using Rubberduck.Inspections;
using Rubberduck.UI;
using Rubberduck.Config;
using Rubberduck.UI.CodeInspections;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;
using Rubberduck.Extensions;

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

            var grammar = GetImplementedSyntax();

            _inspections = GetImplementedCodeInspections();

            EnableCodeInspections();
            var parser = new Parser(grammar);

            _menu = new RubberduckMenu(vbe, addIn, config, parser, _inspections);
            _codeInspectionsToolbar = new CodeInspectionsToolbar(vbe, addIn, parser, _inspections);
        }

        private IList<IInspection> GetImplementedCodeInspections()
                                  {
             var inspections = Assembly.GetExecutingAssembly()
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

             return inspections;
        }

        private static List<ISyntax> GetImplementedSyntax()
        {
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
            return grammar;
        }

        private void EnableCodeInspections()
        {
            foreach (var inspection in _inspections)
            {
                // todo: fetch value from configuration
                
                /*
                foreach (var setting in config.CodeInspectionSettings.CodeInspections)
                {
                    if (inspection.Name == setting.Name)
                    {
                        inspection.Severity = setting.Severity;
                        inspection.InspectionType = settings.InspectionType;
                    }
                } 
                 */
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
