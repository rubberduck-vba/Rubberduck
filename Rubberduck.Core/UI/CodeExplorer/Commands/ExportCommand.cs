using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ExportCommand : CommandBase
    {
        private static readonly Dictionary<ComponentType, string> ExportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ".bas" },
            { ComponentType.ClassModule, ".cls" },
            { ComponentType.Document, ".cls" },
            { ComponentType.UserForm, ".frm" }
        };

        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly IVBE _vbe;

        public ExportCommand(IFileSystemBrowserFactory dialogFactory, IMessageBox messageBox, IProjectsRepository projectsRepository, IVBE vbe)
        {
            _dialogFactory = dialogFactory;
            _vbe = vbe;
            MessageBox = messageBox;
            ProjectsRepository = projectsRepository;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        protected IMessageBox MessageBox { get; }
        protected IProjectsRepository ProjectsRepository { get; }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null)
            {
                return false;
            }

            if (_vbe.Kind != VBEKind.Hosted)
            {
                return false;
            }

            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
            return ExportableFileExtensions.Select(s => s.Key).Contains(componentType);
        }

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null)
            {
                return;
            }

            PromptFileNameAndExport(node.Declaration.QualifiedName.QualifiedModuleName);
        }

        public bool PromptFileNameAndExport(QualifiedModuleName qualifiedModule)
        {
            if (!ExportableFileExtensions.TryGetValue(qualifiedModule.ComponentType, out var extension))
            {
                return false;
            }

            using (var dialog = _dialogFactory.CreateSaveFileDialog())
            {
                dialog.OverwritePrompt = true;
                dialog.FileName = qualifiedModule.ComponentName + extension;

                var result = dialog.ShowDialog();
                if (result != DialogResult.OK)
                {
                    return false;
                }

                var component = ProjectsRepository.Component(qualifiedModule);
                try
                {
                    component.Export(dialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.NotifyWarn(ex.Message, string.Format(Resources.CodeExplorer.CodeExplorerUI.ExportError_Caption, qualifiedModule.ComponentName));
                }                    
                return true;
            }
        }
    }
}
