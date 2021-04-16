using System;
using System.Collections.Generic;
using System.Linq;
using Path = System.IO.Path;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using System.IO.Abstractions;

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
        private readonly IDictionary<ComponentType, List<IRequiredBinaryFilesFromFileNameExtractor>> _binaryFileExtractors;
        private readonly IFileSystem _fileSystem;
        private readonly IConfigurationService<ProjectSettings> _projectSettingsProvider;

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
            IFileSystem fileSystem,
            IMessageBox messageBox,
            IConfigurationService<ProjectSettings> projectSettingsProvider)
            : base(vbeEvents)
        {
            _vbe = vbe;
            _dialogFactory = dialogFactory;
            _parseManager = parseManager;
            _projectsProvider = projectsProvider;
            _moduleNameFromFileExtractor = moduleNameFromFileExtractor;
            _fileSystem = fileSystem;

            _binaryFileExtractors = BinaryFileExtractors(binaryFileExtractors);

            MessageBox = messageBox;
            DeclarationFinderProvider = declarationFinderProvider;

            _projectSettingsProvider = projectSettingsProvider;

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

        private IVBProject TargetProjectFromParameter(object parameter)
        {
            var declaration = (parameter as CodeExplorerItemViewModel)?.Declaration;
            return declaration != null 
                ? _projectsProvider.Project(declaration.ProjectId)
                : null;
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

        private IDictionary<ComponentType, List<IRequiredBinaryFilesFromFileNameExtractor>> BinaryFileExtractors(IEnumerable<IRequiredBinaryFilesFromFileNameExtractor> extractors)
        {
            var dict = new Dictionary<ComponentType, List<IRequiredBinaryFilesFromFileNameExtractor>>();
            foreach (var extractor in extractors)
            {
                foreach (var componentType in extractor.SupportedComponentTypes)
                { 
                    if (!dict.ContainsKey(componentType))
                    {
                        dict.Add(componentType, new List<IRequiredBinaryFilesFromFileNameExtractor>());
                    }

                    dict[componentType].Add(extractor);
                }
            }

            return dict;
        }


        private ProjectSettings _cachedProjectSettings;
        internal void UpdateFilterIndex(int index)
        {
            _cachedProjectSettings.OpenFileDialogFilterIndex = index;
            _projectSettingsProvider.Save(_cachedProjectSettings);
        }

        protected virtual ICollection<string> FilesToImport(object parameter)
        {
            if (_cachedProjectSettings == null)
            {
                _cachedProjectSettings = _projectSettingsProvider.Read();
            }

            using (var dialog = _dialogFactory.CreateOpenFileDialog())
            {
                dialog.AddExtension = true;
                dialog.AutoUpgradeEnabled = true;
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.Multiselect = true;
                dialog.ShowHelp = false;
                dialog.Title = DialogsTitle;
                dialog.FilterIndex = _cachedProjectSettings.OpenFileDialogFilterIndex;
                var vbFilesFilter = IndividualFilter(RubberduckUI.ImportCommand_OpenDialog_Filter_VBFiles, FilterExtension);
                var nonDocumentModulesFilter = IndividualFilter(Resources.CodeExplorer.CodeExplorerUI.ImportCommand_OpenDialog_Filter_NonDocumentModules, AllNonDocumentModulesExtension);
                var standardModulesFilter = IndividualFilter(Resources.CodeExplorer.CodeExplorerUI.ImportCommand_OpenDialog_Filter_StandardModules, StandardModuleExtension);
                var classModulesFilter = IndividualFilter(Resources.CodeExplorer.CodeExplorerUI.ImportCommand_OpenDialog_Filter_ClassModules, ClassModuleExtension);
                var formModulesFilter = IndividualFilter(Resources.CodeExplorer.CodeExplorerUI.ImportCommand_OpenDialog_Filter_FormModules, FormModuleExtension);
                var documentModulesFilter = IndividualFilter(Resources.CodeExplorer.CodeExplorerUI.ImportCommand_OpenDialog_Filter_DocumentModules, DocumentModuleExtension);
                var allFilesFilter = IndividualFilter(RubberduckUI.ImportCommand_OpenDialog_Filter_AllFiles, "*.*");
                dialog.Filter = CompositeFilter("|",
                    vbFilesFilter,
                    nonDocumentModulesFilter,
                    standardModulesFilter,
                    classModulesFilter,
                    formModulesFilter,
                    documentModulesFilter,
                    allFilesFilter);

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return new List<string>();
                }

                if (_cachedProjectSettings.OpenFileDialogFilterIndex != dialog.FilterIndex)
                {
                    UpdateFilterIndex(dialog.FilterIndex);
                }

                var fileNames = dialog.FileNames;
                var fileExtensions = fileNames.Select(Path.GetExtension);
                if (fileExtensions.Any(fileExt => !ImportableExtensions.Contains(fileExt)))
                {
                    NotifyUserAboutAbortDueToUnsupportedFileExtensions(fileNames);
                    return new List<string>();
                }

                return fileNames;
            }

            string IndividualFilter(string filterDescription, string fileExtension)
            {
                return fileExtension != null
                    ? $"{filterDescription} ({fileExtension})|{fileExtension}"
                    : null;
            }

            string CompositeFilter(string separator, params string[] filters)
            {
                return string.Join(separator, filters.Where(filter => !string.IsNullOrEmpty(filter)));
            }
        }

        protected virtual string DialogsTitle => RubberduckUI.ImportCommand_OpenDialog_Title;

        //TODO: Gather all conflicts and report them in one error dialog instead of reporting them one at a time.
        private void NotifyUserAboutAbortDueToUnsupportedFileExtensions(IEnumerable<string> fileNames)
        {
            var firstUnsupportedFile = fileNames.First(filename => !ImportableExtensions.Contains(Path.GetExtension(filename)));
            var unsupportedFileName = Path.GetFileName(firstUnsupportedFile);
            var message = string.Format(RubberduckUI.ImportCommand_UnsupportedFileExtensions, unsupportedFileName);
            MessageBox.NotifyWarn(message, DialogsTitle);
        }

        private void ImportFilesWithSuspension(ICollection<string> filesToImport, IVBProject targetProject)
        {
            var suspendResult = _parseManager.OnSuspendParser(this, new[] {ParserState.Ready}, () => ImportFiles(filesToImport, targetProject));
            var suspendOutcome = suspendResult.Outcome;
            if (suspendOutcome != SuspensionOutcome.Completed)
            {
                if (suspendOutcome == SuspensionOutcome.UnexpectedError || suspendOutcome == SuspensionOutcome.Canceled)
                {
                    //This rethrows the exception with the original stack trace.
                    ExceptionDispatchInfo.Capture(suspendResult.EncounteredException).Throw();
                    return;
                }

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

            if (!ExistingModulesPassPreCheck(existingModules))
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

            if (UserDeclinesExecution(targetProject))
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

        protected virtual bool ExistingModulesPassPreCheck(IDictionary<string, QualifiedModuleName> existingModules) => true;
        protected virtual ICollection<QualifiedModuleName> ModulesToRemoveBeforeImport(IDictionary<string, QualifiedModuleName> existingModules) => new List<QualifiedModuleName>();
        protected virtual bool UserDeclinesExecution(IVBProject targetProject) => false;

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
            if (extension == null || !ComponentTypesForExtension.TryGetValue(extension, out var componentTypes))
            {
                return new List<string>();
            }

            foreach (var componentType in componentTypes)
            {
                if (_binaryFileExtractors.TryGetValue(componentType, out var binaryExtractors))
                {
                    return binaryExtractors
                        .SelectMany(binaryExtractor => binaryExtractor
                            .RequiredBinaryFiles(filename, componentType))
                        .ToHashSet();
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
                    .Where(binaryFileName => !_fileSystem.File.Exists(Path.Combine(path, binaryFileName)))
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


        //We usually only allow extensions to be imported for which we might be able to determine that the conditions are met to actually import the file.
        //However, we ignore this cautionary rule for the extensions specified in AlwaysImportableExtensions;
        protected ICollection<string> ImportableExtensions =>
            ComponentTypesForExtension.Keys
                .Where(fileExtension => ComponentTypesForExtension.TryGetValue(fileExtension, out var componentTypes)
                                        && componentTypes.All(componentType => componentType.BinaryFileExtension() == string.Empty
                                                                               || _binaryFileExtractors.ContainsKey(componentType)
                                                                               || ComponentTypesWithImportMechanismToExistingComponent.Contains(componentType)))
                .Concat(AlwaysImportableExtensions)
                .ToHashSet();

        //TODO: Implement the binary extractors necessary to allow us to remove this member.
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

        private string AllNonDocumentModulesExtension => string.Join("; ",
            ImportableExtensions.Where(ext => !ext.EndsWith("doccls"))
                                .Select(ext => $"*{ext}"));

        private string StandardModuleExtension => "*" + ImportableExtensions.Where(ext => ext.EndsWith("bas")).FirstOrDefault();

        private string ClassModuleExtension => "*" + ImportableExtensions.Where(ext => ext.EndsWith(".cls")).FirstOrDefault();

        private string FormModuleExtension => "*" + ImportableExtensions.Where(ext => ext.EndsWith("frm")).FirstOrDefault();

        private string DocumentModuleExtension => "*" + ImportableExtensions.Where(ext => ext.EndsWith("doccls")).FirstOrDefault();

        protected IDictionary<string, ICollection<ComponentType>> ComponentTypesForExtension { get; }
    }
}