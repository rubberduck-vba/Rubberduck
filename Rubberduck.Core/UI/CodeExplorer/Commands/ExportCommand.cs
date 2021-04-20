using System;
using System.Collections.Generic;
using Path = System.IO.Path;
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
            if (!TryAssignComponentType(parameter, out var componentType))
            {
                return false;
            }

            return _exportableFileExtensions.ContainsKey(componentType);

            bool TryAssignComponentType(object obj, out ComponentType compType)
            {
                if (obj is CodeExplorerComponentViewModel vm)
                {
                    compType = vm.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
                    return true;
                }

                if (obj is CodeExplorerViewModel viewModel)
                {
                    if (viewModel.SelectedItem is CodeExplorerCustomFolderViewModel)
                    {
                        compType = ComponentType.Undefined;
                        return false;
                    }

                    if (viewModel.SelectedItem is CodeExplorerComponentViewModel componentViewModel)
                    {
                        compType = componentViewModel.Declaration.QualifiedName.QualifiedModuleName.ComponentType;
                        return true;
                    }

                    compType = ComponentType.Undefined;
                    return false;
                }

                compType = ComponentType.Undefined;
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var viewModel = (CodeExplorerViewModel)parameter;

            if (!(viewModel.SelectedItem is CodeExplorerComponentViewModel componentViewModel) ||
                componentViewModel.Declaration == null)
            {
                return;
            }

            PromptFileNameAndExport(componentViewModel.Declaration.QualifiedName.QualifiedModuleName, viewModel);
        }

        public bool PromptFileNameAndExport(QualifiedModuleName qualifiedModule)
        {
            return PromptFileNameAndExport(qualifiedModule, null);
        }

        public bool PromptFileNameAndExport(QualifiedModuleName qualifiedModule, CodeExplorerViewModel viewModel)
        {
            if (!_exportableFileExtensions.TryGetValue(qualifiedModule.ComponentType, out var extension))
            {
                return false;
            }

            using (var dialog = _dialogFactory.CreateSaveFileDialog())
            {
                dialog.OverwritePrompt = true;
                dialog.InitialDirectory = viewModel?.CachedExportPath ?? string.Empty;
                dialog.FileName = qualifiedModule.ComponentName + extension;

                var result = dialog.ShowDialog();
                if (result != DialogResult.OK)
                {
                    return false;
                }

                var exportPath = Path.GetDirectoryName(dialog.FileName);
                if (viewModel != null)
                {
                    viewModel.CachedExportPath = exportPath;
                }

                var component = ProjectsProvider.Component(qualifiedModule);
                try
                {
                    var path = Path.GetDirectoryName(dialog.FileName);
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
