using System;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class RemoveCommand : ExportCommand
    {
        public RemoveCommand(IFileSystemBrowserFactory dialogFactory, IMessageBox messageBox, IProjectsProvider projectsProvider) 
            : base(dialogFactory, messageBox, projectsProvider) { }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return base.EvaluateCanExecute(parameter) &&
                   ((CodeExplorerComponentViewModel) parameter).Declaration.QualifiedName.QualifiedModuleName
                   .ComponentType != ComponentType.Document;
        }

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null ||
                node.Declaration.QualifiedName.QualifiedModuleName.ComponentType == ComponentType.Document)
            {
                return;
            }

            var qualifiedModuleName = node.Declaration.QualifiedName.QualifiedModuleName;
            var message = string.Format(CodeExplorerUI.ExportBeforeRemove_Prompt, node.Name);

            switch (MessageBox.Confirm(message, CodeExplorerUI.ExportBeforeRemove_Caption, ConfirmationOutcome.Yes))
            {
                case ConfirmationOutcome.Yes:
                    if (!PromptFileNameAndExport(qualifiedModuleName))
                    {
                        return;
                    }
                    break;
                case ConfirmationOutcome.No:
                    break;
                case ConfirmationOutcome.Cancel:
                    return;
                }

            // No file export or file successfully exported--now remove it
            var components = ProjectsProvider.ComponentsCollection(qualifiedModuleName.ProjectId);
            try
            {
                components?.Remove(ProjectsProvider.Component(qualifiedModuleName));
            }
            catch (Exception ex)
            {
                MessageBox.NotifyWarn(ex.Message, string.Format(CodeExplorerUI.RemoveError_Caption, qualifiedModuleName.ComponentName));
            }
        }
    }
}
