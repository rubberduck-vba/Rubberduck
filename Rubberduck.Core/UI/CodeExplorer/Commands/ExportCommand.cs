using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ExportCommand : CommandBase
    {
        private static readonly Dictionary<ComponentType, string> VBAExportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ComponentTypeExtensions.StandardExtension },
            { ComponentType.ClassModule, ComponentTypeExtensions.ClassExtension },
            { ComponentType.Document, ComponentTypeExtensions.DocClassExtension },
            { ComponentType.UserForm, ComponentTypeExtensions.FormExtension }
        };

        private static readonly Dictionary<ComponentType, string> VB6ExportableFileExtensions = new Dictionary<ComponentType, string>
        {
            { ComponentType.StandardModule, ComponentTypeExtensions.StandardExtension },
            { ComponentType.ClassModule, ComponentTypeExtensions.ClassExtension },
            { ComponentType.VBForm, ComponentTypeExtensions.FormExtension },
            { ComponentType.MDIForm, ComponentTypeExtensions.FormExtension },
            { ComponentType.UserControl, ComponentTypeExtensions.UserControlExtension },
            { ComponentType.DocObject, ComponentTypeExtensions.DocObjectExtension },
            { ComponentType.ActiveXDesigner, ComponentTypeExtensions.ActiveXDesignerExtension },
            { ComponentType.PropPage, ComponentTypeExtensions.PropertyPageExtension },
            { ComponentType.ResFile, ComponentTypeExtensions.ResourceExtension },            
        };

        private readonly IFileSystemBrowserFactory _dialogFactory;
        private readonly Dictionary<ComponentType, string> _exportableFileExtensions;

        public ExportCommand(IFileSystemBrowserFactory dialogFactory, IMessageBox messageBox, IProjectsProvider projectsProvider, IVBE vbe)
        {
            _dialogFactory = dialogFactory;
            MessageBox = messageBox;
            ProjectsProvider = projectsProvider;

            _exportableFileExtensions =
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

            return _exportableFileExtensions.ContainsKey(componentType);
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
            if (!_exportableFileExtensions.TryGetValue(qualifiedModule.ComponentType, out var extension))
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
                    var path = System.IO.Path.GetDirectoryName(dialog.FileName);
                    component.ExportAsSourceFile(path, false, true); // skipped optional parameters interfere with mock setup
                }
                catch (Exception ex)
                {
                    Logger.Warn(ex, $"Failed to export component {qualifiedModule.Name}");
                    MessageBox.NotifyWarn(ex.Message, string.Format(Resources.CodeExplorer.CodeExplorerUI.ExportError_Caption, qualifiedModule.ComponentName));
                }                    
                return true;
            }
        }
    }
}
