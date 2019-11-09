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
        private readonly IDictionary<ComponentType, IRequiredBinaryFilesFromFileNameExtractor> _binaryFileExtractors;
        private readonly IFileExistenceChecker _fileExistenceChecker;

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
            : base(vbe, dialogFactory, vbeEvents, parseManager, messageBox)
        {
            _projectsProvider = projectsProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _moduleNameFromFileExtractor = moduleNameFromFileExtractor;
            _fileExistenceChecker = fileExistenceChecker;

            _binaryFileExtractors = BinaryFileExtractors(binaryFileExtractors);
        }

        private IDictionary<ComponentType, IRequiredBinaryFilesFromFileNameExtractor> BinaryFileExtractors(IEnumerable<IRequiredBinaryFilesFromFileNameExtractor> extractors)
        {
            var dict = new Dictionary<ComponentType, IRequiredBinaryFilesFromFileNameExtractor>();
            foreach (var extractor in extractors)
            {
                foreach (var componentType in extractor.SupportedComponentTypes)
                {
                    if (dict.ContainsKey(componentType))
                    {
                        continue;
                    }

                    dict.Add(componentType, extractor);
                }
            }

            return dict;
        }

        protected override string DialogsTitle => RubberduckUI.UpdateFromFilesCommand_DialogCaption;

        //We only allow extensions to be imported for which we might be able to determine that the conditions are met to actually import the file.
        protected override ICollection<string> ImportableExtensions =>
            base.ImportableExtensions
                .Where(fileExtension => ComponentTypesForExtension.TryGetValue(fileExtension, out var componentTypes) 
                                        && componentTypes.All(componentType => componentType.BinaryFileExtension() == string.Empty 
                                                                               || _binaryFileExtractors.ContainsKey(componentType)
                                                                               || ComponentTypesWithImportMechanismToExistingComponent.Contains(componentType)))
                .ToList();

        //For some component types like user forms and documents we have implemented a way to import them into existing components.
        private ICollection<ComponentType> ComponentTypesWithImportMechanismToExistingComponent => 
            new List<ComponentType>
            {
                ComponentType.Document,
                ComponentType.UserForm
            };

        protected override void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var moduleNames = ModuleNames(filesToImport);

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

            var requiredBinaryFiles = RequiredBinaryFiles(filesToImport);
            var missingBinaries = FilesWithoutRequiredBinaries(requiredBinaryFiles);

            var filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent = FilesWithMechanismToImportToExistingComponent(missingBinaries.Keys);
            var filesWithoutRequiredBinariesWithoutBackupSolution = missingBinaries.Keys
                .Where(fileName => !filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent.Contains(fileName))
                .ToList();

            if (filesWithoutRequiredBinariesWithoutBackupSolution.Any())
            {
                NotifyUserAboutAbortDueToNonExistingBinaryFile(filesWithoutRequiredBinariesWithoutBackupSolution, missingBinaries);
                return;
            }

            if (!filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent.All(filename => modulesToRemoveBeforeImport.ContainsKey(filename)))
            {
                NotifyUserAboutAbortDueToNonExistingBinaryFileAndComponent(
                    filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent,
                    missingBinaries,
                    moduleNames, 
                    modulesToRemoveBeforeImport);
                return;
            }

            //Since we want to import into the existing components, we must not remove them.
            foreach (var filename in filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent)
            {
                modulesToRemoveBeforeImport.Remove(filename);
            }

            var documentFiles = DocumentFiles(moduleNames);

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

            using (var components = targetProject.VBComponents)
            {
                foreach (var filename in filesToImport)
                {
                    if (modulesToRemoveBeforeImport.TryGetValue(filename, out var module))
                    {
                        var component = _projectsProvider.Component(module);
                        components.Remove(component);
                    }

                    if(documentFiles.Contains(filename) || filesWithoutRequiredBinariesWithoutBackupSolution.Contains(filename))
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

        private ICollection<string> DocumentFiles(Dictionary<string, string> moduleNames)
        {
            return moduleNames
                .Select(kvp => kvp.Key)
                .Where(filename => Path.GetExtension(filename) != null
                                   && ComponentTypesForExtension.TryGetValue(Path.GetExtension(filename),
                                       out var componentTypes)
                                   && componentTypes.Contains(ComponentType.Document))
                .ToHashSet();
        }

        private ICollection<string> FilesWithMechanismToImportToExistingComponent(ICollection<string> fileNames)
        {
            return fileNames
                .Where(filename => Path.GetExtension(filename) != null
                                   && ComponentTypesForExtension.TryGetValue(Path.GetExtension(filename), out var componentTypes)
                                   && componentTypes.All(componentType => ComponentTypesWithImportMechanismToExistingComponent.Contains(componentType)))
                .ToHashSet();
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

        private Dictionary<string, ICollection<string>> RequiredBinaryFiles(ICollection<string> fileNames)
        {
            var requiredBinaryNames = new Dictionary<string, ICollection<string>>();
            foreach (var filename in fileNames)
            {
                if (requiredBinaryNames.ContainsKey(filename))
                {
                    continue;
                }

                var requiredBinaryFiles = RequiredBinaryFiles(filename);
                if (requiredBinaryFiles.Any())
                {
                    requiredBinaryNames.Add(filename, requiredBinaryFiles);
                }
            }

            return requiredBinaryNames;
        }

        private ICollection<string> RequiredBinaryFiles(string filename)
        {
            var extension = Path.GetExtension(filename);
            if (!ComponentTypesForExtension.TryGetValue(extension, out var componentTypes))
            {
                return new List<string>();
            }

            foreach (var componentType in componentTypes)
            {
                if (_binaryFileExtractors.TryGetValue(componentType, out var binaryExtractor))
                {
                    return binaryExtractor.RequiredBinaryFiles(filename, componentType);
                }
            }

            return new List<string>();
        }

        private IDictionary<string, ICollection<string>> FilesWithoutRequiredBinaries(Dictionary<string, ICollection<string>> requiredBinaries)
        {
            var filesWithoutBinaries = new Dictionary<string, ICollection<string>>();
            foreach (var (fileName, requiredBinariesForFile) in requiredBinaries)
            {
                var path = Path.GetDirectoryName(fileName);
                var missingBinaries = requiredBinariesForFile
                    .Where(binaryFileName => !_fileExistenceChecker.FileExists(Path.Combine(path, binaryFileName)))
                    .ToList();

                if (missingBinaries.Any())
                {
                    filesWithoutBinaries.Add(fileName, missingBinaries);
                }
            }

            return filesWithoutBinaries;
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
                   && ComponentTypesForExtension.TryGetValue(fileExtension, out var componentTypes) 
                   && componentTypes.Contains(module.ComponentType);
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

        private void NotifyUserAboutAbortDueToNonExistingBinaryFile(ICollection<string> filesWithoutBinary, IDictionary<string, ICollection<string>> missingBinaries)
        {
            var firstFilenameForFileWithoutBinaryAndComponent = filesWithoutBinary.First();
            var missingBinariesOfFirstFilenameWithoutBinaryAndComponent = string.Join(", ", missingBinaries[firstFilenameForFileWithoutBinaryAndComponent]);
            var message = string.Format(
                RubberduckUI.UpdateFromFilesCommand_BinaryDoesNotExist,
                firstFilenameForFileWithoutBinaryAndComponent,
                missingBinariesOfFirstFilenameWithoutBinaryAndComponent);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private void NotifyUserAboutAbortDueToNonExistingBinaryFileAndComponent(ICollection<string> filesWithoutBinary, IDictionary<string, ICollection<string>> missingBinaries, IDictionary<string, string> moduleNames, IDictionary<string, QualifiedModuleName> existingModules)
        {
            var firstFilenameForFileWithoutBinaryAndComponent = filesWithoutBinary.First(filename => !existingModules.ContainsKey(filename));
            var moduleNameOfFirstFilenameWithoutBinaryAndComponent = moduleNames[firstFilenameForFileWithoutBinaryAndComponent];
            var missingBinariesOfFirstFilenameWithoutBinaryAndComponent = string.Join("', '", missingBinaries[firstFilenameForFileWithoutBinaryAndComponent]);
            var message = string.Format(
                RubberduckUI.UpdateFromFilesCommand_BinaryAndComponentDoNotExist,
                firstFilenameForFileWithoutBinaryAndComponent,
                moduleNameOfFirstFilenameWithoutBinaryAndComponent,
                missingBinariesOfFirstFilenameWithoutBinaryAndComponent);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }
    }
}