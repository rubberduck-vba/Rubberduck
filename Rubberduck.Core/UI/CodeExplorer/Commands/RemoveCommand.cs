using System;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
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
            return _exportCommand.CanExecute(parameter) &&
                   parameter is CodeExplorerComponentViewModel viewModel &&
                   viewModel.Declaration.QualifiedName.QualifiedModuleName.ComponentType != ComponentType.Document;
        }

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null ||
                node.Declaration.QualifiedName.QualifiedModuleName.ComponentType == ComponentType.Document)
            {
                return;
            }

            RemoveComponent(node.Declaration.QualifiedModuleName);
        }

        public bool RemoveComponent(QualifiedModuleName qualifiedModuleName, bool promptToExport = true)
        {
            if (promptToExport && !TryExport(qualifiedModuleName))
            {
                return false;
            }
            
            // No file export or file successfully exported--now remove it
            try
            {
                _projectsRepository.RemoveComponent(qualifiedModuleName);
            }
            catch (Exception ex)
            {
                _messageBox.NotifyWarn(ex.Message, string.Format(CodeExplorerUI.RemoveError_Caption, qualifiedModuleName.ComponentName));
                return false;
            }

            return true;
        }

        private bool TryExport(QualifiedModuleName qualifiedModuleName)
        {
            var projectId = qualifiedModuleName.ProjectId;
            var projectType = _projectsRepository.Project(projectId).Type;
            var component = _projectsRepository.Component(qualifiedModuleName);

            if (projectType == ProjectType.HostProject && component.IsSaved)
            {
                return true; // File already up-to-date
            }

            var message = string.Format(CodeExplorerUI.ExportBeforeRemove_Prompt, qualifiedModuleName.Name);

            switch (_messageBox.Confirm(message, CodeExplorerUI.ExportBeforeRemove_Caption, ConfirmationOutcome.Yes))
            {
                case ConfirmationOutcome.No:
                    return true;

                case ConfirmationOutcome.Yes:
                    if (_exportCommand.PromptFileNameAndExport(qualifiedModuleName))
                    {
                        return true;
                    }
                    break;
            }

            return false; // Export cancelled or failed
        }
    }
}
