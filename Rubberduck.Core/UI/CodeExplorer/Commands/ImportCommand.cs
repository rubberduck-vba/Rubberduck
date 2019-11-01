using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
        private readonly IList<string> _importableExtensions;
        private readonly string _filterExtensions;
        private readonly IParseManager _parseManager;

        protected readonly IMessageBox MessageBox;

        public ImportCommand(
            IVBE vbe,
            IFileSystemBrowserFactory dialogFactory,
            IVbeEvents vbeEvents,
            IParseManager parseManager,
            IMessageBox messageBox)
            : base(vbeEvents)
        {
            _vbe = vbe;
            _dialogFactory = dialogFactory;
            _parseManager = parseManager;

            MessageBox = messageBox;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);

            ComponentTypeForExtension = ComponentTypeExtensions.ComponentTypeForExtension(_vbe.Kind);

            _importableExtensions = ComponentTypeForExtension.Keys.ToList();
            _filterExtensions = string.Join("; ", _importableExtensions.Select(ext => $"*{ext}"));

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
                    $"{RubberduckUI.ImportCommand_OpenDialog_Filter_VBFiles} ({_filterExtensions})|{_filterExtensions}|" +
                    $"{RubberduckUI.ImportCommand_OpenDialog_Filter_AllFiles}, (*.*)|*.*";

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return new List<string>();
                }

                var fileNames = dialog.FileNames;
                var fileExtensions = fileNames.Select(Path.GetExtension);
                if (fileExtensions.Any(fileExt => !_importableExtensions.Contains(fileExt)))
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
            var firstUnsupportedFile = fileNames.First(filename => !_importableExtensions.Contains(Path.GetExtension(filename)));
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

        protected virtual void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            using (var components = targetProject.VBComponents)
            {
                foreach (var filename in filesToImport)
                {
                    var fileExtension = Path.GetExtension(filename);
                    if (fileExtension != null
                        && ComponentTypeForExtension.TryGetValue(fileExtension, out var componentType)
                        && componentType == ComponentType.Document)
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

        protected IDictionary<string, ComponentType> ComponentTypeForExtension { get; }
    }
}