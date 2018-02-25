using System.Runtime.InteropServices;
using System.Windows.Input;
using NLog;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.UI.Command
{
    public interface INavigateCommand : ICommand { }

    /// <summary>
    /// A command that navigates to a specified selection, using a <see cref="NavigateCodeEventArgs"/> parameter.
    /// </summary>
    [ComVisible(false)]
    public class NavigateCommand : CommandBase, INavigateCommand
    {
        private readonly IProjectsProvider _projectsProvider;

        public NavigateCommand(IProjectsProvider projectsProvider)
            : base(LogManager.GetCurrentClassLogger())
        {
            _projectsProvider = projectsProvider;
        }

        protected override void OnExecute(object parameter)
        {
            var param = parameter as NavigateCodeEventArgs;
            if(param == null)
            {
                return;
            }

            try
            {
                using (var codeModule = _projectsProvider.Component(param.QualifiedName).CodeModule)
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
