using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenProjectPropertiesCommand : CommandBase
    {
        private readonly IVBE _vbe;

        public OpenProjectPropertiesCommand(IVBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            try
            {
                return parameter != null || _vbe.ProjectsCount == 1;
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            const int openProjectPropertiesId = 2578;

            using (var commandBars = _vbe.CommandBars)
            {
                using (var command = commandBars.FindControl(openProjectPropertiesId))
                {
                    if (_vbe.ProjectsCount == 1)
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

                    var nodeProject = node.GetSelectedDeclaration().Project;
                    if (nodeProject == null)
                    {
                        return; //The project declaration has been disposed, i.e. the project has been removed already.
                    }

                    try
                    {
                        _vbe.ActiveVBProject = nodeProject;
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
}
