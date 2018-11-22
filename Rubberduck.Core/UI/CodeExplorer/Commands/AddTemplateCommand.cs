using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Templates;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddTemplateCommand : CommandBase
    {
        private readonly ITemplateProvider _provider;
        private readonly AddComponentCommand _addComponentCommand;

        public AddTemplateCommand(ITemplateProvider provider, AddComponentCommand addComponentCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _provider = provider;
            _addComponentCommand = addComponentCommand;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            var data = ((string templateName, CodeExplorerItemViewModel model)) parameter;
            return _addComponentCommand.CanAddComponent(data.model,
                new[]
                {
                    ProjectType.HostProject, ProjectType.StandAlone, ProjectType.StandardExe, ProjectType.ActiveXExe
                });
        }

        protected override void OnExecute(object parameter)
        {
            var data = ((string templateName, CodeExplorerItemViewModel model)) parameter;
            if (string.IsNullOrWhiteSpace(data.templateName))
            {
                return;
            }

            var moduleText = GetTemplate(data.templateName);
            _addComponentCommand.AddComponent(data.model, moduleText);        }

        private string GetTemplate(string name)
        {
            return _provider.Load(name).Read();
        }
    }
}