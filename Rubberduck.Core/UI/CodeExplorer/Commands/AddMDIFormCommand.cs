using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class AddMDIFormCommand : AddComponentCommandBase
    {
        private static readonly ProjectType[] Types = { ProjectType.StandardExe, ProjectType.ActiveXExe };

        public AddMDIFormCommand(IVBE vbe) : base(vbe) { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.MDIForm;

        protected override void OnExecute(object parameter)
        {
            AddComponent(parameter as CodeExplorerItemViewModel);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerItemViewModel node))
            {
                return false;
            }

            var project = node.Declaration?.Project;

            if (project != null || Vbe.ProjectsCount != 1)
            {
                return EvaluateCanExecuteForProject(project, node);
            }

            using (var vbProjects = Vbe.VBProjects)
            using (project = vbProjects[1])
            {
                return EvaluateCanExecuteForProject(project, node);
            }
        }

        private bool EvaluateCanExecuteForProject(IVBProject project, CodeExplorerItemViewModel itemViewModel)
        {
            if (project == null)
            {
                return false;
            }

            using (var components = project.VBComponents)
            {
                foreach (var component in components)
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
            }

            return base.EvaluateCanExecute(itemViewModel);
        }
    }
}
