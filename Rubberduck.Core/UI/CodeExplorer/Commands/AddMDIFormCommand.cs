using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddMDIFormCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly AddComponentCommand _addComponentCommand;

        public AddMDIFormCommand(IVBE vbe, AddComponentCommand addComponentCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _addComponentCommand = addComponentCommand;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }
            var project = (node as ICodeExplorerDeclarationViewModel)?.Declaration?.Project;

            if (project == null  && _vbe.ProjectsCount == 1)
            {
                using (var vbProjects = _vbe.VBProjects)
                using (project = vbProjects[1])
                {
                    return EvaluateCanExecuteCore(project, parameter as CodeExplorerItemViewModel);
                }
            }

            return EvaluateCanExecuteCore(project, parameter as CodeExplorerItemViewModel);
        }

        private bool EvaluateCanExecuteCore(IVBProject project, CodeExplorerItemViewModel itemViewModel)
        {
            if (project == null)
            {
                return false;
            }

            foreach (var component in project.VBComponents)
            {
                using (component)
                {
                    if (component.Type == ComponentType.MDIForm)
                    {
                        // Only one MDI Form allowed per project
                        return false;
                    }
                }
            }

            return _addComponentCommand.CanAddComponent(itemViewModel, new[] {ProjectType.StandardExe, ProjectType.ActiveXExe});
        }

        protected override void OnExecute(object parameter)
        {
            _addComponentCommand.AddComponent(parameter as CodeExplorerItemViewModel, ComponentType.MDIForm);
        }
    }
}
