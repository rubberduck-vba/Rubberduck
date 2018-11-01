using System;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.IO;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing;
using Rubberduck.Interaction.Input;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ImportCommand : CommandBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IOpenFileDialog _openFileDialog;
        private readonly IFileHandler _fileHandler;
        private readonly IStringParser _stringParser;

        public ImportCommand(IVBE vbe, IOpenFileDialog openFileDialog, IFileHandler fileHandler, IStringParser stringParser) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _openFileDialog = openFileDialog;
            _fileHandler = fileHandler;
            _stringParser = stringParser;

            _openFileDialog.AddExtension = true;
            _openFileDialog.AutoUpgradeEnabled = true;
            _openFileDialog.CheckFileExists = true;
            _openFileDialog.CheckPathExists = true;
            _openFileDialog.Multiselect = true;
            _openFileDialog.ShowHelp = false;   // we don't want 1996's file picker.
            _openFileDialog.Filter = string.Concat(RubberduckUI.ImportCommand_OpenDialog_Filter_VBFiles, @" (*.cls, *.bas, *.frm, *.doccls)|*.cls; *.bas; *.frm; *.doccls|", RubberduckUI.ImportCommand_OpenDialog_Filter_AllFiles, @" (*.*)|*.*");
            _openFileDialog.Title = RubberduckUI.ImportCommand_OpenDialog_Title;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return parameter != null || _vbe.ProjectsCount == 1 || ThereIsAValidActiveProject();
        }

        private bool ThereIsAValidActiveProject()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                return activeProject != null;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var usingFreshProjectWrapper = false;

            var project = GetNodeProject(parameter as CodeExplorerItemViewModel);

            if (project == null)
            {
                if (_vbe.ProjectsCount == 1)
                {
                    usingFreshProjectWrapper = true;
                    using (var projects = _vbe.VBProjects)
                    {
                        project = projects[1];
                    }
                }
                else if (ThereIsAValidActiveProject())
                {
                    usingFreshProjectWrapper = true;
                    project = _vbe.ActiveVBProject;
                }
            }

            if (project == null || _openFileDialog.ShowDialog() != DialogResult.OK)
            {
                if (usingFreshProjectWrapper)
                {
                    project?.Dispose();
                }
                return;
            }

            var fileExts = _openFileDialog.FileNames.Select(s => s.Split('.').Last());
            if (fileExts.Any(fileExt => !new[] {"bas", "cls", "frm"}.Contains(fileExt)))
            {
                if (usingFreshProjectWrapper)
                {
                    project.Dispose();
                }
                return;
            }

            var uniqueFileCounter = 0;
            foreach (var filename in _openFileDialog.FileNames)
            {
                var sourceText = string.Join(Environment.NewLine, _fileHandler.ReadAllLines(filename));
                var tempHelper = (CodeExplorerItemViewModel)parameter;
                var newFolderName = (parameter is CodeExplorerCustomFolderViewModel) ? tempHelper.Name : tempHelper.GetSelectedDeclaration().CustomFolder;

                var parseResults = _stringParser.Parse(string.Empty, string.Empty, sourceText, new System.Threading.CancellationToken());
                
                var updatedModuleText = FolderAnnotator.AddOrUpdateFolderName(parseResults, newFolderName);
                var extension = Path.GetExtension(filename);
                var importPath = $"RubberduckTempImportFile{uniqueFileCounter++}{extension}";
                try
                {
                    _fileHandler.WriteToFile(importPath, updatedModuleText);

                    using (var components = project.VBComponents)
                    {
                        components.Import(importPath);
                    }
                }
                catch(Exception e)
                {
                    Logger.Error(e); 
                }
                finally
                {
                    if (_fileHandler.Exists(importPath))
                    {
                        _fileHandler.Delete(importPath);
                    }
                }
            }

            if (usingFreshProjectWrapper)
            {
                project.Dispose();
            }
        }

        private IVBProject GetNodeProject(CodeExplorerItemViewModel parameter)
        {
            if (parameter == null)
            {
                return null;
            }

            if (parameter is ICodeExplorerDeclarationViewModel)
            {
                return parameter.GetSelectedDeclaration().Project;
            }

            var node = parameter.Parent;
            while (!(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return ((ICodeExplorerDeclarationViewModel)node).Declaration.Project;
        }

        public void Dispose()
        {
            if (_openFileDialog != null)
            {
                _openFileDialog.Dispose();
            }
        }
    }
}
