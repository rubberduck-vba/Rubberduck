using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class OpenProjectPropertiesCommand : CommandBase
    {
        private readonly VBE _vbe;

        public OpenProjectPropertiesCommand(VBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            try
            {
                using (var projects = _vbe.VBProjects)
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

            using (var projects = _vbe.VBProjects)
            {
                var commandBars = _vbe.CommandBars;
                var command = commandBars.FindControl(Id: openProjectPropertiesId);

                if (projects.Count == 1)
                {
                    command.Execute();
                    Marshal.ReleaseComObject(command);
                    Marshal.ReleaseComObject(commandBars);
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
                    Marshal.ReleaseComObject(command);
                    Marshal.ReleaseComObject(commandBars);
                    return; // the project was probably removed from the VBE, but not from the CE
                }

                command.Execute();
                Marshal.ReleaseComObject(command);
                Marshal.ReleaseComObject(commandBars);
            }
        }
    }
}
