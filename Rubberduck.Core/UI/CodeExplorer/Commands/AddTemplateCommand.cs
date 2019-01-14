using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Templates;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddTemplateCommand : AddComponentCommandBase
    {
        private static readonly ProjectType[] Types = { ProjectType.HostProject, ProjectType.StandAlone, ProjectType.StandardExe, ProjectType.ActiveXExe };

        private readonly ITemplateProvider _provider;

        public AddTemplateCommand(IVBE vbe, ITemplateProvider provider) : base(vbe)
        {
            _provider = provider;
        }

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.Undefined;

        public bool CanExecuteForNode(ICodeExplorerNode model)
        {
            return base.EvaluateCanExecute(model);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            // TODO this cast needs to be safer.
            var data = ((string templateName, CodeExplorerItemViewModel model)) parameter;

            return base.EvaluateCanExecute(data.model);
        }

        protected override void OnExecute(object parameter)
        {
            // TODO this cast needs to be safer.
            var data = ((string templateName, CodeExplorerItemViewModel model)) parameter;

            if (string.IsNullOrWhiteSpace(data.templateName))
            {
                return;
            }

            var moduleText = GetTemplate(data.templateName);
            AddComponent(data.model, moduleText);
        }

        private string GetTemplate(string name)
        {
            return _provider.Load(name).Read();
        }
    }
}