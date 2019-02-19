using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Interaction.Navigation;
using Rubberduck.VBEditor;
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
            : base(LogManager.GetCurrentClassLogger())
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
