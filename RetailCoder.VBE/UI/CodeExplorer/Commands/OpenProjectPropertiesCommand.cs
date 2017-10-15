using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class OpenProjectPropertiesCommand : CommandBase
    {
        private readonly IVBE _vbe;

        public OpenProjectPropertiesCommand(IVBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            try
            {
                var projects = _vbe.VBProjects;
                {
                    return parameter != null || projects.Count == 1;
                }
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            const int openProjectPropertiesId = 2578;

            var projects = _vbe.VBProjects;
            {
                var commandBars = _vbe.CommandBars;
                var command = commandBars.FindControl(openProjectPropertiesId);

                if (projects.Count == 1)
                {
                    command.Execute();
                    return;
                }

                var node = parameter as CodeExplorerItemViewModel;
                while (!(node is ICodeExplorerDeclarationViewModel))
                {
                    // ReSharper disable once PossibleNullReferenceException
                    node = node.Parent; // the project node is an ICodeExplorerDeclarationViewModel--no worries here
                }

                try
                {
                    _vbe.ActiveVBProject = node.GetSelectedDeclaration().Project;
                }
                catch (COMException)
                {
                    return; // the project was probably removed from the VBE, but not from the CE
                }

                command.Execute();
            }
        }
    }
}
