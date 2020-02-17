using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public abstract class AddComponentCommandBase : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly ICodeExplorerAddComponentService _addComponentService;
        private readonly IProjectsProvider _projectsProvider;

        protected AddComponentCommandBase(
            ICodeExplorerAddComponentService addComponentService, 
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider) 
            : base(vbeEvents)
        {
            _addComponentService = addComponentService;
            _projectsProvider = projectsProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        public abstract IEnumerable<ProjectType> AllowableProjectTypes { get; }

        public abstract ComponentType ComponentType { get; }

        protected override void OnExecute(object parameter)
        {
            AddComponent(parameter as CodeExplorerItemViewModel);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerItemViewModel node) 
                || node.Declaration == null)
            {
                return false;
            }

            try
            {
                var project = _projectsProvider.Project(node.Declaration.ProjectId);
                return project != null 
                       && AllowableProjectTypes.Contains(project.Type);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private void AddComponent(CodeExplorerItemViewModel node)
        {
            _addComponentService.AddComponent(node, ComponentType);
        }
    }
}