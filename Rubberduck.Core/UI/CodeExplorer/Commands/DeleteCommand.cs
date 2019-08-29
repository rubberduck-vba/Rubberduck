using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class DeleteCommand : CodeExplorerCommandBase
    {
        private readonly RemoveCommand _removeCommand;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IMessageBox _messageBox;
        private readonly IVBE _vbe;

        public DeleteCommand(
            RemoveCommand removeCommand, 
            IProjectsProvider projectsProvider, 
            IMessageBox messageBox, 
            IVBE vbe, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _removeCommand = removeCommand;
            _projectsProvider = projectsProvider;
            _messageBox = messageBox;
            _vbe = vbe;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes { get; } = new List<Type> { typeof(CodeExplorerComponentViewModel) };

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _vbe.Kind == VBEKind.Standalone &&
                   _removeCommand.CanExecute(parameter);
        }

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) || node.Declaration is null)
            {
                return;
            }

            var qualifiedModuleName = node.Declaration.QualifiedModuleName;
            var component = _projectsProvider.Component(qualifiedModuleName);
            if (component is null)
            {
                return;
            }

            // "{qualifiedModuleName.Name} will be permanently deleted. Continue?" (localized)
            var message = string.Format(Resources.CodeExplorer.CodeExplorerUI.ConfirmBeforeDelete_Prompt, qualifiedModuleName.Name);
            if (!_messageBox.ConfirmYesNo(message, Resources.CodeExplorer.CodeExplorerUI.ConfirmBeforeDelete_Caption))
            {
                return;
            }

            // Have to build the file list *before* removing the component!
            var files = new List<string>();
            for (short i = 1; i <= component.FileCount; i++)
            {
                var fileName = component.GetFileName(i);
                if (fileName != null) // Unsaved components have a null filename for some reason
                {
                    files.Add(component.GetFileName(i));
                }
            }

            if (!_removeCommand.RemoveComponent(qualifiedModuleName, false))
            {
                return; // Remove was cancelled or failed
            }

            var failedDeletions = new List<string>();
            foreach (var file in files)
            {           
                try
                {
                    File.Delete(file);                    
                }
                catch (Exception exception)
                {
                    Logger.Warn(exception, "Failed to delete file");                 
                    failedDeletions.Add(file);
                }
            }

            // Let the user know if there are any component files left on disk
            if (failedDeletions.Any())
            {
                // "The following files could not be deleted: {fileDeletions}" (localized)
                message = string.Format(Resources.CodeExplorer.CodeExplorerUI.DeleteFailed_Message, string.Join(Environment.NewLine, failedDeletions));
                _messageBox.NotifyWarn(message, Resources.CodeExplorer.CodeExplorerUI.DeleteFailed_Caption);
            }
            
        }
    }
}
