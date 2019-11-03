using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class UpdateFromFilesCommand : ImportCommand
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IModuleNameFromFileExtractor _moduleNameFromFileExtractor;

        public UpdateFromFilesCommand(
            IVBE vbe,
            IFileSystemBrowserFactory dialogFactory,
            IVbeEvents vbeEvents,
            IParseManager parseManager,
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider,
            IModuleNameFromFileExtractor moduleNameFromFileExtractor,
            IMessageBox messageBox)
            : base(vbe, dialogFactory, vbeEvents, parseManager, messageBox)
        {
            _projectsProvider = projectsProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _moduleNameFromFileExtractor = moduleNameFromFileExtractor;
        }

        protected override string DialogsTitle => RubberduckUI.UpdateFromFilesCommand_DialogCaption;

        protected override void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var moduleNames = ModuleNames(filesToImport);

            var formBinaryModuleNames = moduleNames
                .Where(kvp => ComponentTypeExtensions.FormBinaryExtension.Equals(Path.GetExtension(kvp.Key)))
                .Select(kvp => kvp.Value)
                .ToHashSet();

            var formFilesWithoutBinaries = FormFilesWithoutBinaries(moduleNames, formBinaryModuleNames);

            //We cannot import the the binary separately.
            foreach (var formBinaryModuleName in formBinaryModuleNames)
            {
                moduleNames.Remove(formBinaryModuleName);
            }

            if (!ValuesAreUnique(moduleNames))
            {
                NotifyUserAboutAbortDueToDuplicateComponent(moduleNames);
                return;
            }

            var modulesToRemoveBeforeImport = Modules(moduleNames, targetProject.ProjectId, finder);

            if(!modulesToRemoveBeforeImport.All(kvp => HasMatchingFileExtension(kvp.Key, kvp.Value)))
            {
                NotifyUserAboutAbortDueToNonMatchingFileExtension(modulesToRemoveBeforeImport);
                return;
            }

            var documentFiles = moduleNames
                .Select(kvp => kvp.Key)
                .Where(filename => Path.GetExtension(filename) != null
                              && ComponentTypeForExtension.TryGetValue(Path.GetExtension(filename), out var componentType)
                              && componentType == ComponentType.Document)
                .ToHashSet();

            //We can only insert into existing documents.
            if (!documentFiles.All(filename => modulesToRemoveBeforeImport.ContainsKey(filename)))
            {
                NotifyUserAboutAbortDueToNonExistingDocument(documentFiles, moduleNames, modulesToRemoveBeforeImport);
                return;
            }

            //We must not remove document modules.
            foreach (var filename in documentFiles)
            {
                modulesToRemoveBeforeImport.Remove(filename);
            }

            //We import the standalone code behind by replacing the code in an existing form.
            //So, the form has to exist already.
            if (!formFilesWithoutBinaries.All(filename => modulesToRemoveBeforeImport.ContainsKey(filename)))
            {
                NotifyUserAboutAbortDueToNonExistingUserForm(documentFiles, moduleNames, modulesToRemoveBeforeImport);
                return;
            }

            foreach (var filename in formFilesWithoutBinaries)
            {
                modulesToRemoveBeforeImport.Remove(filename);
            }

            using (var components = targetProject.VBComponents)
            {
                foreach (var filename in filesToImport)
                {
                    if (modulesToRemoveBeforeImport.TryGetValue(filename, out var module))
                    {
                        var component = _projectsProvider.Component(module);
                        components.Remove(component);
                    }

                    if(documentFiles.Contains(filename) || formBinaryModuleNames.Contains(filename))
                    {
                        //We have to dispose the return value.
                        using (components.ImportSourceFile(filename)) { }
                    }
                    else
                    {
                        //We have to dispose the return value.
                        using (components.Import(filename)) { }
                    }
                }
            }
        }

        private Dictionary<string, string> ModuleNames(ICollection<string> filenames)
        {
            var moduleNames = new Dictionary<string, string>();
            foreach(var filename in filenames)
            {
                if (moduleNames.ContainsKey(filename))
                {
                    continue;
                }

                var moduleName = ModuleName(filename);
                if(moduleName != null)
                {
                    moduleNames.Add(filename, moduleName);
                }
            }

            return moduleNames;
        }

        private string ModuleName(string filename)
        {
            return _moduleNameFromFileExtractor.ModuleName(filename);
        }

        private Dictionary<string, QualifiedModuleName> Modules(IDictionary<string, string> moduleNames, string projectId, DeclarationFinder finder)
        {
            var modules = new Dictionary<string, QualifiedModuleName>();
            foreach (var (fileName, moduleName) in moduleNames)
            {
                var module = Module(moduleName, projectId, finder);
                if (module.HasValue)
                {
                    modules.Add(fileName, module.Value);
                }
            }

            return modules;
        }

        private bool ValuesAreUnique(Dictionary<string, string> moduleNames)
        {
            return moduleNames
                .GroupBy(kvp => kvp.Value)
                .All(moduleNameGroup => moduleNameGroup.Count() == 1);
        }

        private void NotifyUserAboutAbortDueToDuplicateComponent(IDictionary<string, string> moduleNames)
        {
            var firstDuplicateModuleName = moduleNames
                .GroupBy(kvp => kvp.Value)
                .First(moduleNameGroup => moduleNameGroup.Count() > 1)
                .Key;
            var message = string.Format(RubberduckUI.UpdateFromFilesCommand_DuplicateModule, firstDuplicateModuleName);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private ICollection<string> FormFilesWithoutBinaries(IDictionary<string, string> moduleNames, ICollection<string> formBinaryModuleNames)
        {
            return moduleNames
                .Where(kvp => Path.GetExtension(kvp.Key) != null
                              && ComponentTypeForExtension.TryGetValue(Path.GetExtension(kvp.Key), out var componentType)
                              && componentType == ComponentType.UserForm
                              && !formBinaryModuleNames.Contains(kvp.Value))
                .Select(kvp => kvp.Key)
                .ToHashSet();
        }

        private QualifiedModuleName? Module(string moduleName, string projectId, DeclarationFinder finder)
        {
            foreach(var module in finder.AllModules)
            {
                if(module.ProjectId.Equals(projectId)
                    && module.ComponentName.Equals(moduleName))
                {
                    return module;
                }
            }

            return null;
        }

        private bool HasMatchingFileExtension(string filename, QualifiedModuleName module)
        {
            var fileExtension = Path.GetExtension(filename);
            return fileExtension != null 
                   && ComponentTypeForExtension.TryGetValue(fileExtension, out var componentType) 
                   && module.ComponentType.Equals(componentType);
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

        private void NotifyUserAboutAbortDueToNonExistingDocument(ICollection<string> documentFiles, IDictionary<string, string> moduleNames, IDictionary<string, QualifiedModuleName> existingModules)
        {
            var firstNonExistingDocumentFilename = documentFiles.First(filename => !existingModules.ContainsKey(filename));
            var firstNonExistingDocumentModuleName = moduleNames[firstNonExistingDocumentFilename];
            var message = string.Format(
                RubberduckUI.UpdateFromFilesCommand_DocumentDoesNotExist,
                firstNonExistingDocumentModuleName,
                firstNonExistingDocumentFilename);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private void NotifyUserAboutAbortDueToNonExistingUserForm(ICollection<string> userFormFiles, IDictionary<string, string> moduleNames, IDictionary<string, QualifiedModuleName> existingModules)
        {
            var firstNonExistingUserFormFilename = userFormFiles.First(filename => !existingModules.ContainsKey(filename));
            var firstNonExistingUserFormModuleName = moduleNames[firstNonExistingUserFormFilename];
            var message = string.Format(
                RubberduckUI.UpdateFromFilesCommand_UserFormDoesNotExist,
                firstNonExistingUserFormModuleName,
                firstNonExistingUserFormFilename);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }
    }
}