using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class AddMDIFormCommand : AddComponentCommandBase
    {
        private static readonly ProjectType[] Types = { ProjectType.StandardExe, ProjectType.ActiveXExe };

        private readonly IProjectsProvider _projectsProvider;

        public AddMDIFormCommand(
            ICodeExplorerAddComponentService addComponentService, 
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider) 
            : base(addComponentService, vbeEvents, projectsProvider)
        {
            _projectsProvider = projectsProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.MDIForm;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerItemViewModel node))
            {
                return false;
            }

            var project = node.Declaration != null
                ? _projectsProvider.Project(node.Declaration.ProjectId)
                : null;

            return EvaluateCanExecuteForProject(project);
        }

        private static bool EvaluateCanExecuteForProject(IVBProject project)
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

            return true;
        }
    }
}
