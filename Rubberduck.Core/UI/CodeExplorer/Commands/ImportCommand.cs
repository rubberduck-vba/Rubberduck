using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Events;
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

        public ImportCommand(IVBE vbe, IFileSystemBrowserFactory dialogFactory, IVbeEvents vbeEvents) : base(vbeEvents)
        {
            _vbe = vbe;
            _dialogFactory = dialogFactory;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);

            _importableExtensions =
                vbe.Kind == VBEKind.Hosted
                    ? new List<string> {"bas", "cls", "frm", "doccls"} // VBA 
                    : new List<string> {"bas", "cls", "frm", "ctl", "pag", "dob"}; // VB6

            _filterExtensions = string.Join("; ", _importableExtensions.Select(ext => $"*.{ext}"));

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
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

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
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
                dialog.AddExtension = true;
                dialog.AutoUpgradeEnabled = true;
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.Multiselect = true;
                dialog.ShowHelp = false;
                dialog.Title = RubberduckUI.ImportCommand_OpenDialog_Title;
                dialog.Filter = 
                    $"{RubberduckUI.ImportCommand_OpenDialog_Filter_VBFiles} ({_filterExtensions})|{_filterExtensions}|" +
                    $"{RubberduckUI.ImportCommand_OpenDialog_Filter_AllFiles}, (*.*)|*.*";

                if (project == null || dialog.ShowDialog() != DialogResult.OK)
                {
                    if (usingFreshProjectWrapper)
                    {
                        project?.Dispose();
                    }
                    return;
                }

                var fileExists = dialog.FileNames.Select(s => s.Split('.').Last());
                if (fileExists.Any(fileExt => !_importableExtensions.Contains(fileExt)))
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
    }
}