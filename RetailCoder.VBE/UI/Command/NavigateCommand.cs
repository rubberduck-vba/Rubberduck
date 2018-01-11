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

        protected override void OnExecute(object parameter)
        {
            var param = parameter as NavigateCodeEventArgs;
            if (param == null || param.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                using (var codeModule = param.QualifiedName.Component.CodeModule)
                {
                    using (var codePane = codeModule.CodePane)
                    {
                        codePane.Selection = param.Selection;
                    }
                }
            }
            catch (COMException)
            {
            }
        }
    }
}
