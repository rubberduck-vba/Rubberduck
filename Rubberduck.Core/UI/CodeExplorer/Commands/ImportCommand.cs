using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
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

        public ImportCommand(IVBE vbe, IFileSystemBrowserFactory dialogFactory)
        {
            _vbe = vbe;
            _dialogFactory = dialogFactory;
        }

        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            return base.EvaluateCanExecute(parameter) && _vbe.ProjectsCount == 1 || ThereIsAValidActiveProject();
        }

        private bool ThereIsAValidActiveProject()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                return activeProject != null;
            }
        }

        private static readonly List<string> ImportableExtensions = new List<string> { "bas", "cls", "frm" };

        protected override void OnExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter))
            {
                return;
            }

            var usingFreshProjectWrapper = false;
            var project = (parameter as CodeExplorerItemViewModel)?.Declaration?.Project;

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
                else
                {
                    return;
                }
            }

            using (var dialog = _dialogFactory.CreateOpenFileDialog())
            {
                ConfigureOpenDialog(dialog);

                if (project == null || dialog.ShowDialog() != DialogResult.OK)
                {
                    if (usingFreshProjectWrapper)
                    {
                        project?.Dispose();
                    }
                    return;
                }

                var fileExists = dialog.FileNames.Select(s => s.Split('.').Last());
                if (fileExists.Any(fileExt => !ImportableExtensions.Contains(fileExt)))
                {
                    if (usingFreshProjectWrapper)
                    {
                        project.Dispose();
                    }
                    return;
                }

                foreach (var filename in dialog.FileNames)
                {
                    using (var components = project.VBComponents)
                    {
                        components.Import(filename);
                    }
                }
            }

            if (usingFreshProjectWrapper)
            {
                project.Dispose();
            }
        }

        private static void ConfigureOpenDialog(IOpenFileDialog dialog)
        {
            dialog.AddExtension = true;
            dialog.AutoUpgradeEnabled = true;
            dialog.CheckFileExists = true;
            dialog.CheckPathExists = true;
            dialog.Multiselect = true;
            dialog.ShowHelp = false;   // we don't want 1996's file picker.
            //TODO - Filter needs descriptions.
            dialog.Filter = string.Concat(RubberduckUI.ImportCommand_OpenDialog_Filter_VBFiles,
                @" (*.cls, *.bas, *.frm, *.doccls)|*.cls; *.bas; *.frm; *.doccls|",
                RubberduckUI.ImportCommand_OpenDialog_Filter_AllFiles, @" (*.*)|*.*");
            dialog.Title = RubberduckUI.ImportCommand_OpenDialog_Title;
        }
    }
}