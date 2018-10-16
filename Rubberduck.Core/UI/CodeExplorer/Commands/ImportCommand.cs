using System;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.IO;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ImportCommand : CommandBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IOpenFileDialog _openFileDialog;

        public ImportCommand(IVBE vbe, IOpenFileDialog openFileDialog) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _openFileDialog = openFileDialog;

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

            foreach (var filename in _openFileDialog.FileNames)
            {
                var extension = filename.Split('.').Last();

                var sourceText = string.Join(Environment.NewLine, File.ReadAllLines(filename));

                string updatedModuleText;

                var result = VBACodeStringParser.Parse(sourceText, t => t.startRule());

                var tempHelper = (CodeExplorerItemViewModel)parameter;
                var updatedFolder = (parameter is CodeExplorerCustomFolderViewModel) ? tempHelper.Name : tempHelper.GetSelectedDeclaration().CustomFolder;

                const string folderAnnotation = "'@Folder";
                if (result.parseTree.GetChild(0).GetText().Contains(folderAnnotation))
                {
                    var workingTree = TreeContainingFolderAnnotation(result.parseTree, folderAnnotation);

                    var originalAnnotation = workingTree.GetText();

                    var originalFolder = FolderNameFromFolderAnnotation(workingTree, folderAnnotation);

                    var updatedFolderAnnotation = originalAnnotation.Replace(originalFolder, updatedFolder);
                    result.rewriter.Replace(workingTree.SourceInterval.a, workingTree.SourceInterval.b, updatedFolderAnnotation);

                    updatedModuleText = result.rewriter.GetText();
                }
                else
                {
                    updatedModuleText = $"{folderAnnotation}({updatedFolder}){Environment.NewLine}" + result.rewriter.GetText();
                }

                try
                {
                    var tempFile = $"RubberduckTempImportFile.{extension}";
                    var sw = File.CreateText(tempFile);
                    sw.Write(updatedModuleText);
                    sw.Close();

                    using (var components = project.VBComponents)
                    {
                        components.Import(tempFile);
                    }

                    File.Delete(tempFile);
                }
                catch
                {
                    Logger.Log(LogLevel.Error, "Unable to create temporary file to import into the correct folder while expecuting " + nameof(ImportCommand));
                }
            }

            if (usingFreshProjectWrapper)
            {
                project.Dispose();
            }
        }

        private string FolderNameFromFolderAnnotation(IParseTree parseTree, string folderAnnotation)
        {
            var searchAnnotation = parseTree.GetText();
            var folderNameEnclosedInQuotes = searchAnnotation.Contains('"');
            var enclosingCharacter = folderNameEnclosedInQuotes ? '"' : ')';
            var startIndex = searchAnnotation.IndexOf(folderAnnotation) + 1 + folderAnnotation.Length + (folderNameEnclosedInQuotes ? 1 : 0);
            int endIndex = searchAnnotation.IndexOf(enclosingCharacter, startIndex + 1);
            var length = endIndex - startIndex;

            return searchAnnotation.Substring(startIndex, length);
        }

        private IParseTree TreeContainingFolderAnnotation(IParseTree containingTree, string folderAnnotation)
        {
            for (int i=0;i< containingTree.ChildCount; i++)
            {
                if (containingTree.GetChild(i).GetText().Contains(folderAnnotation))
                {
                    return TreeContainingFolderAnnotation(containingTree.GetChild(i), folderAnnotation);
                }
            }

            return containingTree;
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
