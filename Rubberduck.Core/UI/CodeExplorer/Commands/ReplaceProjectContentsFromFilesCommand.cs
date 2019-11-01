using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ReplaceProjectContentsFromFilesCommand : ImportCommand
    {
        public ReplaceProjectContentsFromFilesCommand(
            IVBE vbe, 
            IFileSystemBrowserFactory dialogFactory, 
            IVbeEvents vbeEvents, 
            IParseManager parseManager) 
            :base(vbe, dialogFactory, vbeEvents, parseManager)
        {}

        protected override void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            //TODO: Ask for confirmation to delete the project contents and replace them with the selected modules.

            RemoveReimportableComponents(targetProject);
            base.ImportFiles(filesToImport, targetProject);
        }

        private void RemoveReimportableComponents(IVBProject project)
        {
            var importableComponentTypes = ComponentTypeForExtension.Values;
            using(var components = project.VBComponents)
            {
                foreach(var component in components)
                {
                    using (component)
                    {
                        if (importableComponentTypes.Contains(component.Type))
                        {
                            components.Remove(component);
                        }
                    }
                }
            }
        }
    }
}
