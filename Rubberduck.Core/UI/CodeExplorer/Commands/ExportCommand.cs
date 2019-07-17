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
        private static readonly Dictionary<ComponentType, string> VBAExportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ".bas" },
            { ComponentType.ClassModule, ".cls" },
            { ComponentType.Document, ".cls" },
            { ComponentType.UserForm, ".frm" }            
        };

        private static readonly Dictionary<ComponentType, string> VB6ExportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ".bas" },
            { ComponentType.ClassModule, ".cls" },
            { ComponentType.VBForm, ".frm" },
            { ComponentType.MDIForm, ".frm" },
            { ComponentType.UserControl, ".ctl" },
            { ComponentType.DocObject, ".dob" },
            { ComponentType.ActiveXDesigner, ".dsr" },
            { ComponentType.PropPage, ".pag" },
            { ComponentType.ResFile, ".res" },            
        };

        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly Dictionary<ComponentType, string> _ExportableFileExtensions;

        public ExportCommand(IFileSystemBrowserFactory dialogFactory, IMessageBox messageBox, IProjectsProvider projectsProvider, IVBE vbe)
        {
            _dialogFactory = dialogFactory;
            MessageBox = messageBox;
            ProjectsProvider = projectsProvider;

            _ExportableFileExtensions =
                vbe.Kind == VBEKind.Hosted
                    ? VBAExportableFileExtensions
                    : VB6ExportableFileExtensions;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        protected IMessageBox MessageBox { get; }
        protected IProjectsProvider ProjectsProvider { get; }
        
        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (!(parameter is CodeExplorerComponentViewModel node) ||
                node.Declaration == null)
            {
                return false;
            }

            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.ComponentType;

            return _ExportableFileExtensions.Select(s => s.Key).Contains(componentType);
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
            if (!_ExportableFileExtensions.TryGetValue(qualifiedModule.ComponentType, out var extension))
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
