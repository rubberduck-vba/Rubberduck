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
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing;

using Rubberduck.Parsing;
using Antlr4.Runtime;

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
                var extension = Path.GetExtension(filename);

                var sourceText = string.Join(Environment.NewLine, File.ReadAllLines(filename));
                var tempHelper = (CodeExplorerItemViewModel)parameter;
                var newFolderName = (parameter is CodeExplorerCustomFolderViewModel) ? tempHelper.Name : tempHelper.GetSelectedDeclaration().CustomFolder;

                var startRule = VBACodeStringParser.Parse(sourceText, t => t.startRule());

                #region Delete This Old Code
                //if (HasModuleDeclarations(startRule.parseTree, out var moduleDeclarations))
                //{
                //    if (HasFolderAnnotation(startRule.parseTree, out var folderAnnotation))
                //    {
                //        var oldFolder = folderAnnotation.GetChild<VBAParser.AnnotationArgListContext>()
                //            .GetChild<VBAParser.AnnotationArgContext>();
                //        startRule.rewriter.Replace(oldFolder.SourceInterval.a, oldFolder.SourceInterval.b, newFolderName);
                //    }
                //    else
                //    {
                //        if (HasOptionExplicit(startRule.parseTree, out var optionExplicitStmt))
                //        {
                //            startRule.rewriter.InsertBefore(optionExplicitStmt.SourceInterval.a, FolderAnnotationWithFolderName(newFolderName) + Environment.NewLine);
                //        }
                //        else
                //        {
                //            var lastNewline = moduleDeclarations.GetDescendents<VBAParser.EndOfLineContext>().Last();
                //            startRule.rewriter.InsertBefore(lastNewline.SourceInterval.a, FolderAnnotationWithFolderName(newFolderName) + Environment.NewLine);
                //        }
                //    }
                //}
                //else
                //{
                //    startRule.rewriter.InsertBefore(startRule.parseTree.GetChild(0).SourceInterval.a,
                //        FolderAnnotationWithFolderName(newFolderName) + Environment.NewLine + Environment.NewLine);
                //}

                //var updatedModuleText = startRule.rewriter.GetText();
                #endregion

                var updatedModuleText = RewrittenModuleText(startRule, newFolderName);
                try
                {
                    var tempFile = $"RubberduckTempImportFile{extension}";
                    var sw = File.CreateText(tempFile);
                    sw.Write(updatedModuleText);
                    sw.Close();

                    using (var components = project.VBComponents)
                    {
                        components.Import(tempFile);
                    }

                    File.Delete(tempFile);
                }
                catch(Exception e)
                {
                    Logger.Error(e); 
                }
            }

            if (usingFreshProjectWrapper)
            {
                project.Dispose();
            }
        }

        private string RewrittenModuleText((IParseTree parseTree, TokenStreamRewriter rewriter) startRule, string updatedFolderName)
        {
            if (HasModuleDeclarations(startRule.parseTree, out var moduleDeclarations))
            {
                if (HasFolderAnnotation(startRule.parseTree, out var folderAnnotation))
                {
                    var oldFolder = folderAnnotation.GetChild<VBAParser.AnnotationArgListContext>()
                        .GetChild<VBAParser.AnnotationArgContext>();
                    startRule.rewriter.Replace(oldFolder.SourceInterval.a, oldFolder.SourceInterval.b, updatedFolderName);
                }
                else
                {
                    if (HasOptionExplicit(startRule.parseTree, out var optionExplicitStmt))
                    {
                        startRule.rewriter.InsertBefore(optionExplicitStmt.SourceInterval.a, FolderAnnotationWithFolderName(updatedFolderName) + Environment.NewLine);
                    }
                    else
                    {
                        var lastNewline = moduleDeclarations.GetDescendents<VBAParser.EndOfLineContext>().Last();
                        startRule.rewriter.InsertBefore(lastNewline.SourceInterval.a, FolderAnnotationWithFolderName(updatedFolderName) + Environment.NewLine);
                    }
                }
            }
            else
            {
                startRule.rewriter.InsertBefore(startRule.parseTree.GetChild(0).SourceInterval.a,
                    FolderAnnotationWithFolderName(updatedFolderName) + Environment.NewLine + Environment.NewLine);
            }

            return startRule.rewriter.GetText();
        }

        private bool HasModuleDeclarations(IParseTree parseTree, out VBAParser.ModuleDeclarationsContext moduleDeclarations)
        {
            var startRuleContext = (ParserRuleContext)parseTree;
            var moduleDescendents = startRuleContext.GetDescendents<VBAParser.ModuleDeclarationsContext>();
            if (!moduleDescendents.ElementAt(0).GetText().Equals(string.Empty))
            {
                moduleDeclarations = moduleDescendents.ElementAt(0);
                return true;
            }

            moduleDeclarations = null;
            return false;
        }

        private bool HasFolderAnnotation(IParseTree parseTree, out VBAParser.AnnotationContext folderAnnotation)
        {
            var startRuleContext = (ParserRuleContext)parseTree;
            var folderDescendents = startRuleContext.GetDescendents<VBAParser.AnnotationContext>()
                                        .Where(a => a.GetText().Contains(AnnotationType.Folder.ToString()));
            if (folderDescendents.Count() > 0)
            {
                folderAnnotation = folderDescendents.ElementAt(0);
                return true;
            }

            folderAnnotation = null;
            return false;
        }

        private bool HasOptionExplicit(IParseTree parseTree, out VBAParser.OptionExplicitStmtContext optionExplicit)
        {
            var startRuleContext = (ParserRuleContext)parseTree;
            var optionExplicitDescendents = startRuleContext.GetDescendents<VBAParser.OptionExplicitStmtContext>();
            if (optionExplicitDescendents.Count() > 0)
            {
                optionExplicit = optionExplicitDescendents.ElementAt(0);
                return true;
            }

            optionExplicit = null;
            return false;
        }

        private string FolderAnnotationWithFolderName(string folderName)
        {
            return $"'@{AnnotationType.Folder}({folderName})";
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
