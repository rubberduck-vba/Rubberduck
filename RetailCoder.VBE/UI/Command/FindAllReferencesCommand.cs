using System;
using Rubberduck.Navigation;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Command
{
    public class FindAllReferencesCommand : ICommand<Declaration>
    {
        private readonly IDeclarationNavigator _service;

        public FindAllReferencesCommand([FindReferences] IDeclarationNavigator service)
        {
            _service = service;
        }

        public void Execute(Declaration parameter)
        {
            _service.Find(parameter);
        }

        public void Execute()
        {
            _service.Find();
        }
    }

    [AttributeUsage(AttributeTargets.Parameter)]
    public class FindReferencesAttribute : Attribute
    {
    }
}