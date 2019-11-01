using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
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
