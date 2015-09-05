using System.Runtime.InteropServices;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that navigates to a specified selection, using a <see cref="NavigateCodeEventArgs"/> parameter.
    /// </summary>
    [ComVisible(false)]
    public class NavigateCommand : CommandBase
    {
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public NavigateCommand(ICodePaneWrapperFactory wrapperFactory)
        {
            _wrapperFactory = wrapperFactory;
        }

        public override void Execute(object parameter)
        {
            var param = parameter as NavigateCodeEventArgs;
            if (param == null || param.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                var codePane = _wrapperFactory.Create(param.QualifiedName.Component.CodeModule.CodePane);
                codePane.Selection = param.Selection;
            }
            catch (COMException)
            {
            }
        }
    }
}
