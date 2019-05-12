using System.Runtime.InteropServices;
using Rubberduck.Interaction.Navigation;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command
{

    /// <summary>
    /// A command that navigates to a specified selection, using a <see cref="NavigateCodeEventArgs"/> parameter.
    /// </summary>
    [ComVisible(false)]
    public class NavigateCommand : CommandBase, INavigateCommand
    {
        private readonly ISelectionService _selectionService;

        public NavigateCommand(ISelectionService selectionService)
        {
            _selectionService = selectionService;
        }

        protected override void OnExecute(object parameter)
        {
            var param = parameter as NavigateCodeEventArgs;
            if(param == null)
            {
                return;
            }

            _selectionService.TrySetActiveSelection(param.QualifiedName, param.Selection);
        }
    }
}
