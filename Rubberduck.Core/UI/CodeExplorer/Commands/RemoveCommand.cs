using System;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class RemoveCommand : CommandBase
    {
        private readonly ExportCommand _exportCommand;
        private readonly IProjectsRepository _projectsRepository;
        private readonly IMessageBox _messageBox;
        private readonly IVBE _vbe;

        public RemoveCommand(ExportCommand exportCommand, IProjectsRepository projectsRepository, IMessageBox messageBox, IVBE vbe) 
        {
            _exportCommand = exportCommand;
            _projectsRepository = projectsRepository;
            _messageBox = messageBox;
            _vbe = vbe;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _vbe.Kind == VBEKind.Standalone ||  
                   _exportCommand.CanExecute(parameter) &&
                   ((CodeExplorerComponentViewModel)parameter).Declaration.QualifiedName.QualifiedModuleName.ComponentType != ComponentType.Document;
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
            var projectId = qualifiedModuleName.ProjectId;
            var projectType = _projectsRepository.Project(projectId).Type;

            if (projectType == ProjectType.HostProject) // Prompt for export for VBA only
            {
                var message = string.Format(CodeExplorerUI.ExportBeforeRemove_Prompt, node.Name);

                switch(_messageBox.Confirm(message, CodeExplorerUI.ExportBeforeRemove_Caption, ConfirmationOutcome.Yes))
                {
                    case ConfirmationOutcome.Yes:
                        if (!_exportCommand.PromptFileNameAndExport(qualifiedModuleName))
                        {
                            return; // Don't remove if export was cancelled
                        }
                        break;

                    case ConfirmationOutcome.Cancel:
                        return;

                    case ConfirmationOutcome.No:
                        break;
                }
            }

            // No file export or file successfully exported--now remove it
            try
            {
                _projectsRepository.RemoveComponent(qualifiedModuleName);
            }
            catch (Exception ex)
            {
                _messageBox.NotifyWarn(ex.Message, string.Format(CodeExplorerUI.RemoveError_Caption, qualifiedModuleName.ComponentName));
            }
        }
    }
}
