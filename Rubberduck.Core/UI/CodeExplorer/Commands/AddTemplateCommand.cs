using System;
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
            if (parameter is null)
            {
                return false;
            }

            try
            {
                // TODO this cast needs to be safer.
                var data = ((string templateName, ICodeExplorerNode model))parameter;

                return base.EvaluateCanExecute(data.model);
            }
            catch (Exception ex)
            {
                Logger.Trace(ex);
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
                // TODO this cast needs to be safer.
                var data = ((string templateName, ICodeExplorerNode node))parameter;

                if (string.IsNullOrWhiteSpace(data.templateName) || !(data.node is CodeExplorerItemViewModel model))
                {
                    return;
                }

                var moduleText = GetTemplate(data.templateName);
                AddComponent(model, moduleText);
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