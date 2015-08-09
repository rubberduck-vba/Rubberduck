using System.Runtime.InteropServices;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command
{
    public class NavigateCommand : ICommand<NavigateCodeEventArgs>
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public NavigateCommand(ICodePaneWrapperFactory wrapperFactory)
        {
            _wrapperFactory = wrapperFactory;
        }

        public void Execute(NavigateCodeEventArgs parameter)
        {
            if (parameter == null || parameter.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                var codePane = _wrapperFactory.Create(parameter.QualifiedName.Component.CodeModule.CodePane);
                codePane.Selection = parameter.Selection;
            }
            catch (COMException)
            {
            }
        }

        public void Execute()
        {
            Execute(null);
        }
    }
}
