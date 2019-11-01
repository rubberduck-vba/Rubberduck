using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ReplaceProjectContentsFromFilesCommand : ImportCommand
    {
        public ReplaceProjectContentsFromFilesCommand(
            IVBE vbe, 
            IFileSystemBrowserFactory dialogFactory, 
            IVbeEvents vbeEvents, 
            IParseManager parseManager,
            IMessageBox messageBox) 
            :base(vbe, dialogFactory, vbeEvents, parseManager, messageBox)
        { }

        protected override string DialogsTitle => RubberduckUI.ReplaceProjectContentsFromFilesCommand_DialogCaption;

        protected override void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            if (!UserConfirmsToReplaceProjectContents(targetProject))
            {
                return;
            }

            RemoveReimportableComponents(targetProject);
            base.ImportFiles(filesToImport, targetProject);
        }

        private bool UserConfirmsToReplaceProjectContents(IVBProject project)
        {
            var projectName = project.Name;
            var message = string.Format(RubberduckUI.ReplaceProjectContentsFromFilesCommand_DialogCaption, projectName);
            return MessageBox.ConfirmYesNo(message, DialogsTitle, false);
        }

        private void RemoveReimportableComponents(IVBProject project)
        {
            var reimportableComponentTypes = ComponentTypeForExtension.Values
                .Where(componentType => componentType != ComponentType.Document);
            using(var components = project.VBComponents)
            {
                foreach(var component in components)
                {
                    using (component)
                    {
                        if (reimportableComponentTypes.Contains(component.Type))
                        {
                            components.Remove(component);
                        }
                    }
                }
            }
        }
    }
}
