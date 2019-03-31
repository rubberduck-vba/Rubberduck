using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

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

        public ExportCommand(IFileSystemBrowserFactory dialogFactory, IMessageBox messageBox, IProjectsProvider projectsProvider)
            : base(LogManager.GetCurrentClassLogger())
        {
            _dialogFactory = dialogFactory;
            MessageBox = messageBox;
            ProjectsProvider = projectsProvider;
        }

        protected IMessageBox MessageBox { get; }
        protected IProjectsProvider ProjectsProvider { get; }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null)
            {
                return false;
            }

            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
            return ExportableFileExtensions.Select(s => s.Key).Contains(componentType);
        }

        protected override void OnExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter) || 
                !(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null)
            {
                return;
            }

            PromptFileNameAndExport(node.Declaration.QualifiedName.QualifiedModuleName);
        }

        protected bool PromptFileNameAndExport(QualifiedModuleName qualifiedModule)
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

                var component = ProjectsProvider.Component(qualifiedModule);
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
