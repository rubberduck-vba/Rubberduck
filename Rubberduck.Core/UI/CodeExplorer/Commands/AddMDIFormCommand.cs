using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
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
                project = _vbe.VBProjects[1];
            }

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

            return _addComponentCommand.CanAddComponent(parameter as CodeExplorerItemViewModel, new[] { ProjectType.StandardExe, ProjectType.ActiveXExe });
        }

        protected override void OnExecute(object parameter)
        {
            _addComponentCommand.AddComponent(parameter as CodeExplorerItemViewModel, ComponentType.MDIForm);
        }
    }
}
