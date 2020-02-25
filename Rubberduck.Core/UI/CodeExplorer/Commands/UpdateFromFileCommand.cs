using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class UpdateFromFilesCommand : ImportCommand
    {
        public UpdateFromFilesCommand(
            IVBE vbe,
            IFileSystemBrowserFactory dialogFactory,
            IVbeEvents vbeEvents,
            IParseManager parseManager,
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider,
            IModuleNameFromFileExtractor moduleNameFromFileExtractor,
            IEnumerable<IRequiredBinaryFilesFromFileNameExtractor> binaryFileExtractors,
            IFileExistenceChecker fileExistenceChecker,
            IMessageBox messageBox)
            : base(vbe, dialogFactory, vbeEvents, parseManager, declarationFinderProvider, projectsProvider, moduleNameFromFileExtractor, binaryFileExtractors, fileExistenceChecker, messageBox)
        {}

        protected override string DialogsTitle => RubberduckUI.UpdateFromFilesCommand_DialogCaption;

        //Since we remove the components, we keep on the safe side.
        protected override IEnumerable<string> AlwaysImportableExtensions => Enumerable.Empty<string>();

        protected override bool ExistingModulesPassPreCheck(IDictionary<string, QualifiedModuleName> existingModules)
        {
            if (!existingModules.All(kvp => HasMatchingFileExtension(kvp.Key, kvp.Value)))
            {
                NotifyUserAboutAbortDueToNonMatchingFileExtension(existingModules);
                return false;
            }

            return true;
        }

        protected override ICollection<QualifiedModuleName> ModulesToRemoveBeforeImport(IDictionary<string, QualifiedModuleName> existingModules)
        {
            return existingModules.Values.ToList();
        }

        private void NotifyUserAboutAbortDueToNonMatchingFileExtension(IDictionary<string, QualifiedModuleName> modules)
        {
            var (firstNonMatchingFileName, firstNonMatchingModule) = modules.First(kvp => !HasMatchingFileExtension(kvp.Key, kvp.Value));
            var message = string.Format(
                RubberduckUI.UpdateFromFilesCommand_DifferentComponentType,
                firstNonMatchingModule.ComponentName, 
                firstNonMatchingFileName);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }
    }
}