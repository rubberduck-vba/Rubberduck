using System;
using Rubberduck.Navigation;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Command
{
    public class FindAllImplementationsCommand : ICommand<Declaration>
    {
        private readonly IDeclarationNavigator _service;

        public FindAllImplementationsCommand([FindImplementations] IDeclarationNavigator service)
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
    public class FindImplementationsAttribute : Attribute
    {
    }

    public class FindAllImplementationsCommandMenuItem : CommandMenuItemBase
    {
        public FindAllImplementationsCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "ContextMenu_GoToImplementation"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindImplementations; } }
    }
}