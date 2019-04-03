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

        public AddTemplateCommand(IVBE vbe, ITemplateProvider provider) 
            : base(vbe)
        {
            _provider = provider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.Undefined;

        public bool CanExecuteForNode(ICodeExplorerNode model)
        {
            return CanExecute(model);
        }

        //FIXME: This evaluate method assumes a parameter that is incompatible with the contract of the base class. This causes it to always return false.
        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (parameter is null)
            {
                return false;
            }

            if (parameter is ICodeExplorerNode)
            {
                return true;
            }

            try
            {
                if(parameter is System.ValueTuple<string, ICodeExplorerNode> data)
                {
                    return CanExecute(data.Item2);
                }

                return false;
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