using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Templates;
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

        public AddTemplateCommand(ICodeExplorerAddComponentService addComponentService, ITemplateProvider provider) 
        {
            _provider = provider;
            _addComponentService = addComponentService;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        public IEnumerable<ProjectType> AllowableProjectTypes => ApplicableProjectTypes;

        public ComponentType ComponentType => ComponentType.Undefined;

        public bool CanExecuteForNode(ICodeExplorerNode model)
        {
            return CanExecute(model);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            try
            {
                if(parameter is System.ValueTuple<string, ICodeExplorerNode> data)
                {
                    return EvaluateCanExecute(data.Item2);
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.Trace(ex);
                return false;
            }
        }

        private bool EvaluateCanExecute(ICodeExplorerNode node)
        {
            if (!(node is CodeExplorerItemViewModel)
                || node.Declaration == null)
            {
                return false;
            }

            try
            {
                var project = node.Declaration.Project;
                return AllowableProjectTypes.Contains(project.Type);
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is null)
            {
                return;
            }

            try
            {
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
            catch (Exception ex)
            {
                Logger.Trace(ex);
            }
        }

        private string GetTemplate(string name)
        {
            return _provider.Load(name).Read();
        }
    }
}