using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
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

            RemoveReImportableComponents(targetProject);
            base.ImportFiles(filesToImport, targetProject);
        }

        private bool UserConfirmsToReplaceProjectContents(IVBProject project)
        {
            var projectName = project.Name;
            var message = string.Format(RubberduckUI.ReplaceProjectContentsFromFilesCommand_DialogCaption, projectName);
            return MessageBox.ConfirmYesNo(message, DialogsTitle, false);
        }

        private void RemoveReImportableComponents(IVBProject project)
        {
            var reImportableComponentTypes = ReImportableComponentTypes;
            using(var components = project.VBComponents)
            {
                foreach(var component in components)
                {
                    using (component)
                    {
                        if (reImportableComponentTypes.Contains(component.Type))
                        {
                            components.Remove(component);
                        }
                    }
                }
            }
        }

        //We currently do not take precautions for component types requiring a binary file to be present.
        private ICollection<ComponentType> ReImportableComponentTypes => ComponentTypesForExtension.Values
            .SelectMany(componentTypes => componentTypes)
            .Where(componentType => componentType != ComponentType.Document)
            .ToList();
    }
}
