using System;
using System.Runtime.InteropServices;
using Rubberduck.Navigation;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that finds all implementations of a specified method, or of the active interface module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllImplementationsCommand : CommandBase
    {
        private readonly IDeclarationNavigator _service;

        public FindAllImplementationsCommand([FindImplementations] IDeclarationNavigator service)
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
    public class FindImplementationsAttribute : Attribute
    {
    }
}