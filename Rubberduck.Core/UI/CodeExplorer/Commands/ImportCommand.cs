using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ImportCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IVBE _vbe;
        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly IParseManager _parseManager;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IModuleNameFromFileExtractor _moduleNameFromFileExtractor;
        private readonly IDictionary<ComponentType, IRequiredBinaryFilesFromFileNameExtractor> _binaryFileExtractors;
        private readonly IFileExistenceChecker _fileExistenceChecker;

        protected readonly IDeclarationFinderProvider DeclarationFinderProvider;
        protected readonly IMessageBox MessageBox;

        public ImportCommand(
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
            : base(vbeEvents)
        {
            _vbe = vbe;
            _dialogFactory = dialogFactory;
            _parseManager = parseManager;
            _projectsProvider = projectsProvider;
            DeclarationFinderProvider = declarationFinderProvider;
            _moduleNameFromFileExtractor = moduleNameFromFileExtractor;
            _fileExistenceChecker = fileExistenceChecker;

            _binaryFileExtractors = BinaryFileExtractors(binaryFileExtractors);

            MessageBox = messageBox;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);

            ComponentTypesForExtension = ComponentTypeExtensions.ComponentTypesForExtension(_vbe.Kind);

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
            AddToOnExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _vbe.ProjectsCount == 1 || ThereIsAValidActiveProject();
        }

        private bool ThereIsAValidActiveProject()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                return activeProject != null;
            }
        }

        private (IVBProject project, bool needsDisposal) TargetProject(object parameter)
        {
            var targetProject = TargetProjectFromParameter(parameter);
            if (targetProject != null)
            {
                return (targetProject, false);
            }

            targetProject = TargetProjectFromVbe();

            return (targetProject, targetProject != null);
        }

        private static IVBProject TargetProjectFromParameter(object parameter)
        {
            return (parameter as CodeExplorerItemViewModel)?.Declaration?.Project;
        }

        private IVBProject TargetProjectFromVbe()
        {
            if (_vbe.ProjectsCount == 1)
            {
                using (var projects = _vbe.VBProjects)
                {
                    return projects[1];
                }
            }

            var activeProject = _vbe.ActiveVBProject;
            return activeProject != null && !activeProject.IsWrappingNullReference
                ? activeProject
                : null;
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

        protected virtual ICollection<string> FilesToImport(object parameter)
        {
            using (var dialog = _dialogFactory.CreateOpenFileDialog())
            {
                dialog.AddExtension = true;
                dialog.AutoUpgradeEnabled = true;
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.Multiselect = true;
                dialog.ShowHelp = false;
                dialog.Title = DialogsTitle;
                dialog.Filter =
                    $"{RubberduckUI.ImportCommand_OpenDialog_Filter_VBFiles} ({FilterExtension})|{FilterExtension}|" +
                    $"{RubberduckUI.ImportCommand_OpenDialog_Filter_AllFiles}, (*.*)|*.*";

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return new List<string>();
                }

                var fileNames = dialog.FileNames;
                var fileExtensions = fileNames.Select(Path.GetExtension);
                var importableExtensions = ImportableExtensions;
                if (fileExtensions.Any(fileExt => !importableExtensions.Contains(fileExt)))
                {
                    NotifyUserAboutAbortDueToUnsupportedFileExtensions(fileNames);
                    return new List<string>();
                }

                return fileNames;
            }
        }

        protected virtual string DialogsTitle => RubberduckUI.ImportCommand_OpenDialog_Title;

        private void NotifyUserAboutAbortDueToUnsupportedFileExtensions(IEnumerable<string> fileNames)
        {
            var firstUnsupportedFile = fileNames.First(filename => !ImportableExtensions.Contains(Path.GetExtension(filename)));
            var unsupportedFileName = Path.GetFileName(firstUnsupportedFile);
            var message = string.Format(RubberduckUI.ImportCommand_UnsupportedFileExtensions, unsupportedFileName);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private void ImportFilesWithSuspension(ICollection<string> filesToImport, IVBProject targetProject)
        {
            var suspensionResult = _parseManager.OnSuspendParser(this, new[] {ParserState.Ready}, () => ImportFiles(filesToImport, targetProject));
            if (suspensionResult != SuspensionResult.Completed)
            {
                Logger.Warn("File import failed due to suspension failure.");
            }
        }

        protected void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;

            var moduleNames = ModuleNames(filesToImport);

             if (!ValuesAreUnique(moduleNames))
            {
                NotifyUserAboutAbortDueToDuplicateComponent(moduleNames);
                return;
            }

            var existingModules = Modules(moduleNames, targetProject.ProjectId, finder);

            if (!ExistingModulesAreGenerallyOk(existingModules))
            {
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

            if (!filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent.All(filename => existingModules.ContainsKey(filename) 
                                                                                                         && HasMatchingFileExtension(filename, existingModules[filename])))
            {
                NotifyUserAboutAbortDueToNonExistingBinaryFileAndComponent(
                    filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent,
                    missingBinaries,
                    moduleNames,
                    existingModules);
                return;
            }

            var modulesToRemoveBeforeImport = ModulesToRemoveBeforeImport(existingModules);

            //Since we want to import into the existing components, we must not remove them.
            foreach (var filename in filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent)
            {
                var module = existingModules[filename];
                if (modulesToRemoveBeforeImport.Contains(module))
                {
                    modulesToRemoveBeforeImport.Remove(module);
                }
            }

            var documentFiles = DocumentFiles(moduleNames);

            //We can only insert into existing documents.
            if (!documentFiles.All(filename => existingModules.ContainsKey(filename)
                                               && HasMatchingFileExtension(filename, existingModules[filename])))
            {
                NotifyUserAboutAbortDueToNonExistingDocument(documentFiles, moduleNames, existingModules);
                return;
            }

            //We must not remove component types we cannot reimport. modules.
            var reImportableComponentTypes = ReImportableComponentTypes;
            modulesToRemoveBeforeImport = modulesToRemoveBeforeImport
                .Where(module => reImportableComponentTypes.Contains(module.ComponentType))
                .ToList();

            if (UserDeniesExecution(targetProject))
            {
                return;
            }

            using (var components = targetProject.VBComponents)
            {
                foreach (var module in modulesToRemoveBeforeImport)
                {
                    var component = _projectsProvider.Component(module);
                    components.Remove(component);
                }

                foreach (var filename in filesToImport)
                {
                    if (documentFiles.Contains(filename) || filesWithoutRequiredBinaryButWithPossibilityToImportToExistingComponent.Contains(filename))
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

        protected virtual bool ExistingModulesAreGenerallyOk(IDictionary<string, QualifiedModuleName> existingModules) => true;
        protected virtual ICollection<QualifiedModuleName> ModulesToRemoveBeforeImport(IDictionary<string, QualifiedModuleName> existingModules) => new List<QualifiedModuleName>();
        protected virtual bool UserDeniesExecution(IVBProject targetProject) => false;

        protected bool HasMatchingFileExtension(string filename, QualifiedModuleName module)
        {
            var fileExtension = Path.GetExtension(filename);
            return fileExtension != null
                   && ComponentTypesForExtension.TryGetValue(fileExtension, out var componentTypes)
                   && componentTypes.Contains(module.ComponentType);
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
            foreach (var filename in filenames)
            {
                if (moduleNames.ContainsKey(filename))
                {
                    continue;
                }

                var moduleName = ModuleName(filename);
                if (moduleName != null)
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
            var message = string.Format(RubberduckUI.ImportCommand_DuplicateModule, firstDuplicateModuleName);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private QualifiedModuleName? Module(string moduleName, string projectId, DeclarationFinder finder)
        {
            foreach (var module in finder.AllModules)
            {
                if (module.ProjectId.Equals(projectId)
                    && module.ComponentName.Equals(moduleName))
                {
                    return module;
                }
            }

            return null;
        }

        private void NotifyUserAboutAbortDueToNonExistingDocument(ICollection<string> documentFiles, IDictionary<string, string> moduleNames, IDictionary<string, QualifiedModuleName> existingModules)
        {
            var firstNonExistingDocumentFilename = documentFiles.First(filename => !existingModules.ContainsKey(filename));
            var firstNonExistingDocumentModuleName = moduleNames[firstNonExistingDocumentFilename];
            var message = string.Format(
                RubberduckUI.ImportCommand_DocumentDoesNotExist,
                firstNonExistingDocumentModuleName,
                firstNonExistingDocumentFilename);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private void NotifyUserAboutAbortDueToNonExistingBinaryFile(ICollection<string> filesWithoutBinary, IDictionary<string, ICollection<string>> missingBinaries)
        {
            var firstFilenameForFileWithoutBinaryAndComponent = filesWithoutBinary.First();
            var missingBinariesOfFirstFilenameWithoutBinaryAndComponent = string.Join(", ", missingBinaries[firstFilenameForFileWithoutBinaryAndComponent]);
            var message = string.Format(
                RubberduckUI.ImportCommand_BinaryDoesNotExist,
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
                RubberduckUI.ImportCommand_BinaryAndComponentDoNotExist,
                firstFilenameForFileWithoutBinaryAndComponent,
                moduleNameOfFirstFilenameWithoutBinaryAndComponent,
                missingBinariesOfFirstFilenameWithoutBinaryAndComponent);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        protected override void OnExecute(object parameter)
        {
            var (targetProject, targetProjectNeedsDisposal) = TargetProject(parameter);

            if (targetProject == null)
            {
                return;
            }

            var filesToImport = FilesToImport(parameter);

            if (!filesToImport.Any())
            {
                return;
            }

            ImportFilesWithSuspension(filesToImport, targetProject);

            if (targetProjectNeedsDisposal)
            {
                targetProject.Dispose();
            }
        }


        //We only allow extensions to be imported for which we might be able to determine that the conditions are met to actually import the file.
        //The exception are specif exceptions to the rule.
        protected ICollection<string> ImportableExtensions =>
            ComponentTypesForExtension.Keys
                .Where(fileExtension => ComponentTypesForExtension.TryGetValue(fileExtension, out var componentTypes)
                                        && componentTypes.All(componentType => componentType.BinaryFileExtension() == string.Empty
                                                                               || _binaryFileExtractors.ContainsKey(componentType)
                                                                               || ComponentTypesWithImportMechanismToExistingComponent.Contains(componentType)))
                .Concat(AlwaysImportableExtensions)
                .ToHashSet();

        protected virtual IEnumerable<string> AlwaysImportableExtensions => _vbe.Kind == VBEKind.Standalone
            ? ComponentTypesForExtension.Keys
            : Enumerable.Empty<string>();

        //For some component types like user forms and documents we have implemented a way to import them into existing components.
        private static ICollection<ComponentType> ComponentTypesWithImportMechanismToExistingComponent =>
            new List<ComponentType>
            {
                ComponentType.Document,
                ComponentType.UserForm
            };

        private ICollection<ComponentType> ReImportableComponentTypes => ComponentTypesForExtension.Values
            .SelectMany(componentTypes => componentTypes)
            .Where(componentType => componentType != ComponentType.Document)
            .ToList();

        private string FilterExtension => string.Join("; ", ImportableExtensions.Select(ext => $"*{ext}"));

        protected IDictionary<string, ICollection<ComponentType>> ComponentTypesForExtension { get; }
    }
}