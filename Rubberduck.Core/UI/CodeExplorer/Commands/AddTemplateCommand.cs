using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Templates;
using Rubberduck.VBEditor.Events;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddTemplateCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private static readonly ProjectType[] ApplicableProjectTypes =
        {
            ProjectType.HostProject,
            ProjectType.StandAlone,
            ProjectType.StandardExe,
            ProjectType.ActiveXExe
        };

        private readonly ITemplateProvider _provider;
        private readonly ICodeExplorerAddComponentService _addComponentService;
        private readonly IProjectsProvider _projectsProvider;

        public AddTemplateCommand(
                ICodeExplorerAddComponentService addComponentService, 
                ITemplateProvider provider, 
                IVbeEvents vbeEvents,
                IProjectsProvider projectsProvider) 
                : base(vbeEvents)
        {
            _provider = provider;
            _addComponentService = addComponentService;
            _projectsProvider = projectsProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => new[]{typeof(System.ValueTuple<string, ICodeExplorerNode>)};

        public IEnumerable<ProjectType> AllowableProjectTypes => ApplicableProjectTypes;

        //We need a valid component type to add the component in the first place. Then the module content gets overwritten.
        //TODO: Find a way to pass in the correct component type for a template. (A wrong component type does not hurt in VBA, but in VB6 it does.)
        public ComponentType ComponentType => ComponentType.ClassModule;

        public bool CanExecuteForNode(ICodeExplorerNode model)
        {
            return EvaluateCanExecute(model);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if(parameter is System.ValueTuple<string, ICodeExplorerNode> data)
            {
                return EvaluateCanExecute(data.Item2);
            }

            return false;
        }

        private bool EvaluateCanExecute(ICodeExplorerNode node)
        {
            if (!ApplicableNodes.Contains(node.GetType())
                || !(node is CodeExplorerItemViewModel)
                || node.Declaration == null)
            {
                return false;
            }

            var project = _projectsProvider.Project(node.Declaration.ProjectId);
            return project != null 
                   && AllowableProjectTypes.Contains(project.Type);
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is null)
            {
                return;
            }

            if (!(parameter is System.ValueTuple<string, ICodeExplorerNode> data))
            {
                return;
            }

            var (templateName, node) = data;

            if (string.IsNullOrWhiteSpace(templateName) || !(node is CodeExplorerItemViewModel model))
            {
                return;
            }

            var moduleText = GetTemplate(templateName);
            _addComponentService.AddComponentWithAttributes(model, ComponentType, moduleText);
        }

        private string GetTemplate(string name)
        {
            return _provider.Load(name).Read();
        }
    }
}