using System.Runtime.InteropServices;
using System.Windows.Input;
using NLog;

namespace Rubberduck.UI.Command
{
    public interface INavigateCommand : ICommand { }

    /// <summary>
    /// A command that navigates to a specified selection, using a <see cref="NavigateCodeEventArgs"/> parameter.
    /// </summary>
    [ComVisible(false)]
    public class NavigateCommand : CommandBase, INavigateCommand
    {
        public NavigateCommand() : base(LogManager.GetCurrentClassLogger()) { }

        protected override void ExecuteImpl(object parameter)
        {
            var param = parameter as NavigateCodeEventArgs;
            if (param == null || param.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                param.QualifiedName.Component.CodeModule.CodePane.Selection = param.Selection;
            }
            catch (COMException)
            {
            }
        }
    }
}
