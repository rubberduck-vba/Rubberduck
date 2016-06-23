using System.Runtime.InteropServices;
using System.Windows.Input;

namespace Rubberduck.UI.Command
{
    public interface INavigateCommand : ICommand { }

    /// <summary>
    /// A command that navigates to a specified selection, using a <see cref="NavigateCodeEventArgs"/> parameter.
    /// </summary>
    [ComVisible(false)]
    public class NavigateCommand : CommandBase, INavigateCommand
    {
        public override void Execute(object parameter)
        {
            var param = parameter as NavigateCodeEventArgs;
            if (param == null || param.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                var pane = param.QualifiedName.Component.CodeModule.CodePane;
                var selection = param.Selection;

                pane.SetSelection(selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn);
                pane.ForceFocus();
            }
            catch (COMException)
            {
            }
        }
    }
}
