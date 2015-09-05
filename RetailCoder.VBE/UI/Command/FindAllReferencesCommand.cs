using System;
using System.Runtime.InteropServices;
using Rubberduck.Navigation;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllReferencesCommand : CommandBase
    {
        private readonly IDeclarationNavigator _service;

        public FindAllReferencesCommand([FindReferences] IDeclarationNavigator service)
        {
            _service = service;
        }

        public override void Execute(object parameter)
        {
            if (parameter == null)
            {
                _service.Find();
                return;
            }

            var param = (Declaration)parameter;
            _service.Find(param);
        }

    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public class FindReferencesAttribute : Attribute
    {
    }
}