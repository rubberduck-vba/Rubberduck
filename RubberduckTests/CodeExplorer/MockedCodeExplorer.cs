﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.UIContext;
using Rubberduck.SettingsProvider;
using Rubberduck.Interaction;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.UnitTesting.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Settings;

namespace RubberduckTests.CodeExplorer
{
    internal class MockedCodeExplorer : IDisposable
    {
        private readonly GeneralSettings _generalSettings = new GeneralSettings();

        private readonly Mock<IUiDispatcher> _uiDispatcher = new Mock<IUiDispatcher>();
        private readonly Mock<IConfigurationService<GeneralSettings>> _generalSettingsProvider = new Mock<IConfigurationService<GeneralSettings>>();
        private readonly Mock<IConfigurationService<WindowSettings>> _windowSettingsProvider = new Mock<IConfigurationService<WindowSettings>>();
        private readonly Mock<IConfigurationService<UnitTestSettings>> _unitTestSettingsProvider = new Mock<IConfigurationService<UnitTestSettings>>();
        private readonly Mock<ConfigurationLoader> _configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
        private readonly Mock<IVBEInteraction> _interaction = new Mock<IVBEInteraction>();

        public MockedCodeExplorer()
        {
            _generalSettingsProvider.Setup(s => s.Read()).Returns(_generalSettings);
            _windowSettingsProvider.Setup(s => s.Read()).Returns(WindowSettings);
            _configLoader.Setup(c => c.Read()).Returns(GetDefaultUnitTestConfig());
            _unitTestSettingsProvider.Setup(s => s.Read())
                .Returns(new UnitTestSettings(BindingMode.LateBinding,AssertMode.StrictAssert, true, true, false));

            _uiDispatcher.Setup(m => m.Invoke(It.IsAny<Action>())).Callback((Action argument) => argument.Invoke());
            _uiDispatcher.Setup(m => m.StartTask(It.IsAny<Action>(), It.IsAny<TaskCreationOptions>())).Returns((Action argument, TaskCreationOptions options) => Task.Factory.StartNew(argument.Invoke, options));

            SaveDialog = new Mock<ISaveFileDialog>();
            SaveDialog.Setup(o => o.OverwritePrompt);

            OpenDialog = new Mock<IOpenFileDialog>();
            OpenDialog.Setup(o => o.AddExtension);
            OpenDialog.Setup(o => o.AutoUpgradeEnabled);
            OpenDialog.Setup(o => o.CheckFileExists);
            OpenDialog.Setup(o => o.Multiselect);
            OpenDialog.Setup(o => o.ShowHelp);
            OpenDialog.Setup(o => o.Filter);
            OpenDialog.Setup(o => o.CheckFileExists);

            FolderBrowser = new Mock<IFolderBrowser>();

            BrowserFactory = new Mock<IFileSystemBrowserFactory>();
            BrowserFactory.Setup(m => m.CreateSaveFileDialog()).Returns(SaveDialog.Object);
            BrowserFactory.Setup(m => m.CreateOpenFileDialog()).Returns(OpenDialog.Object);
            BrowserFactory
                .Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true,
                    @"C:\Users\Rubberduck\Documents\Subfolder")).Returns(FolderBrowser.Object);
        }

        public MockedCodeExplorer(string code) : this(ProjectType.HostProject, ComponentType.StandardModule, code) { }

        public MockedCodeExplorer(ProjectType projectType, ComponentType componentType = ComponentType.StandardModule, string code = "") : this()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("TestModule", componentType, code);

            VbComponents = project.MockVBComponents;
            VbComponent = project.MockComponents.First();
            VbProject = project.Build();
            Vbe = builder.AddProject(VbProject).Build();

            ProjectsRepository = new Mock<IProjectsRepository>();
            ProjectsRepository.Setup(x => x.Project(It.IsAny<string>())).Returns(VbProject.Object);
            ProjectsRepository.Setup(x => x.Component(It.IsAny<QualifiedModuleName>())).Returns(VbComponent.Object);

            SetupViewModelAndParse();
        }

        public MockedCodeExplorer(ProjectType projectType,
            IReadOnlyList<ComponentType> componentTypes,
            IReadOnlyList<string> code = null) : this()
        {
            if (code != null && componentTypes.Count != code.Count)
            {
                Assert.Inconclusive("MockedCodeExplorer Setup Error");
            }

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType);

            for (var index = 0; index < componentTypes.Count; index++)
            {
                var item = componentTypes[index];
                if (item == ComponentType.UserForm)
                {
                    project.MockUserFormBuilder($"{item.ToString()}{index}", code is null ? string.Empty : code[index]).AddFormToProjectBuilder();
                }
                else
                {
                    project.AddComponent($"{item.ToString()}{index}", item, code is null ? string.Empty : code[index]);
                }
            }

            VbComponents = project.MockVBComponents;
            VbComponent = project.MockComponents.First();
            VbProject = project.Build();
            Vbe = builder.AddProject(VbProject).Build();

            ProjectsRepository = new Mock<IProjectsRepository>();
            ProjectsRepository.Setup(x => x.Project(It.IsAny<string>())).Returns(VbProject.Object);
            ProjectsRepository.Setup(x => x.Component(It.IsAny<QualifiedModuleName>())).Returns(VbComponent.Object);

            SetupViewModelAndParse();

            VbProject.SetupGet(m => m.VBComponents.Count).Returns(componentTypes.Count);
        }

        private void SetupViewModelAndParse()
        {
            var parser = MockParser.Create(Vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(Vbe));
            State = parser.State;

            var exportCommand = new ExportCommand(BrowserFactory.Object, MessageBox.Object, State.ProjectsProvider, Vbe.Object);
            var removeCommand = new RemoveCommand(exportCommand, ProjectsRepository.Object, MessageBox.Object, Vbe.Object);

            ViewModel = new CodeExplorerViewModel(State, removeCommand,
                _generalSettingsProvider.Object,
                _windowSettingsProvider.Object,
                _uiDispatcher.Object, Vbe.Object,
                null,
                new CodeExplorerSyncProvider(Vbe.Object, State));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
        }

        public RubberduckParserState State { get; set; }
        public Mock<IVBE> Vbe { get; }
        public CodeExplorerViewModel ViewModel { get; set; }
        public Mock<IVBProject> VbProject { get; }
        public Mock<IVBComponents> VbComponents { get; }
        public Mock<IVBComponent> VbComponent { get; }
        public Mock<IFileSystemBrowserFactory> BrowserFactory { get; }
        public Mock<ISaveFileDialog> SaveDialog { get; }
        public Mock<IOpenFileDialog> OpenDialog { get; }
        public Mock<IFolderBrowser> FolderBrowser { get; }
        public Mock<IMessageBox> MessageBox { get; } = new Mock<IMessageBox>();
        public Mock<IProjectsRepository> ProjectsRepository { get; }

        public WindowSettings WindowSettings { get; } = new WindowSettings();

        public MockedCodeExplorer ImplementAddStdModuleCommand()
        {
            ViewModel.AddStdModuleCommand = new AddStdModuleCommand(AddComponentService());
            return this;
        }

        private ICodeExplorerAddComponentService AddComponentService()
        {
            var codePaneComponentSourceCodeHandler = new CodeModuleComponentSourceCodeHandler();

            var addComponentBaseService = new AddComponentService(State.ProjectsProvider, codePaneComponentSourceCodeHandler, codePaneComponentSourceCodeHandler);

            return new CodeExplorerAddComponentService(State, addComponentBaseService, Vbe.Object);
        }

        public void ExecuteAddStdModuleCommand()
        {
            if (ViewModel.AddStdModuleCommand is null)
            {
                ImplementAddStdModuleCommand();
            }
            ViewModel.AddStdModuleCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddClassModuleCommand()
        {
            ViewModel.AddClassModuleCommand = new AddClassModuleCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddClassModuleCommand()
        {
            if (ViewModel.AddClassModuleCommand is null)
            {
                ImplementAddClassModuleCommand();
            }
            ViewModel.AddClassModuleCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddUserFormCommand()
        {
            ViewModel.AddUserFormCommand = new AddUserFormCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddUserFormCommand()
        {
            if (ViewModel.AddUserFormCommand is null)
            {
                ImplementAddUserFormCommand();
            }
            ViewModel.AddUserFormCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddVbFormCommand()
        {
            ViewModel.AddVBFormCommand = new AddVBFormCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddVbFormCommand()
        {
            if (ViewModel.AddVBFormCommand is null)
            {
                ImplementAddVbFormCommand();
            }
            ViewModel.AddVBFormCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddMdiFormCommand()
        {
            ViewModel.AddMDIFormCommand = new AddMDIFormCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddMdiFormCommand()
        {
            if (ViewModel.AddMDIFormCommand is null)
            {
                ImplementAddMdiFormCommand();
            }
            ViewModel.AddMDIFormCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddUserControlCommand()
        {
            ViewModel.AddUserControlCommand = new AddUserControlCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddUserControlCommand()
        {
            if (ViewModel.AddUserControlCommand is null)
            {
                ImplementAddUserControlCommand();
            }
            ViewModel.AddUserControlCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddPropertyPageCommand()
        {
            ViewModel.AddPropertyPageCommand = new AddPropertyPageCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddPropertyPageCommand()
        {
            if (ViewModel.AddPropertyPageCommand is null)
            {
                ImplementAddPropertyPageCommand();
            }
            ViewModel.AddPropertyPageCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddUserDocumentCommand()
        {
            ViewModel.AddUserDocumentCommand = new AddUserDocumentCommand(AddComponentService());
            return this;
        }

        public void ExecuteAddUserDocumentCommand()
        {
            if (ViewModel.AddUserDocumentCommand is null)
            {
                ImplementAddUserDocumentCommand();
            }
            ViewModel.AddUserDocumentCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddTestModuleCommand()
        {
            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var codeGenerator = new TestCodeGenerator(Vbe.Object, State, MessageBox.Object, _interaction.Object, _unitTestSettingsProvider.Object, indenter, null);
            ViewModel.AddTestModuleCommand = new AddTestComponentCommand(Vbe.Object, State, codeGenerator);
            return this;
        }

        public void ExecuteAddTestModuleCommand()
        {
            if (ViewModel.AddTestModuleCommand is null)
            {
                ImplementAddTestModuleCommand();
            }
            ViewModel.AddTestModuleCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementAddTestModuleWithStubsCommand()
        {
            ImplementAddTestModuleCommand();
            ViewModel.AddTestModuleWithStubsCommand = new AddTestModuleWithStubsCommand(Vbe.Object, ViewModel.AddTestModuleCommand);
            return this;
        }

        public void ExecuteAddTestModuleWithStubsCommand()
        {
            if (ViewModel.AddTestModuleWithStubsCommand is null)
            {
                ImplementAddTestModuleWithStubsCommand();
            }
            ViewModel.AddTestModuleWithStubsCommand.Execute(ViewModel.SelectedItem);
        }

        public void ExecuteImportCommand()
        {
            ViewModel.ImportCommand = new ImportCommand(Vbe.Object, BrowserFactory.Object);
            ViewModel.ImportCommand.Execute(ViewModel.SelectedItem);
        }

        public void ExecuteExportAllCommand()
        {
            if (ViewModel.ExportAllCommand is null)
            {
                ImplementExportAllCommand();
            }
            ViewModel.ExportAllCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementExportAllCommand()
        {
            ViewModel.ExportAllCommand = new ExportAllCommand(Vbe.Object, BrowserFactory.Object);
            return this;
        }

        public void ExecuteExportCommand()
        {
            if (ViewModel.ExportCommand is null)
            {
                ImplementExportCommand();
            }
            ViewModel.ExportCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementExportCommand()
        {
            ViewModel.ExportCommand = new ExportCommand(BrowserFactory.Object, MessageBox.Object, State.ProjectsProvider, Vbe.Object);
            return this;
        }

        public void ExecuteOpenDesignerCommand()
        {
            if (ViewModel.OpenDesignerCommand is null)
            {
                ImplementOpenDesignerCommand();
            }
            ViewModel.OpenDesignerCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementOpenDesignerCommand()
        {
            ViewModel.OpenDesignerCommand = new OpenDesignerCommand(State.ProjectsProvider, Vbe.Object);
            return this;
        }

        public void ExecuteIndenterCommand()
        {
            if (ViewModel.IndenterCommand is null)
            {
                ImplementIndenterCommand();
            }
            ViewModel.IndenterCommand.Execute(ViewModel.SelectedItem);
        }

        public MockedCodeExplorer ImplementIndenterCommand()
        {
            ViewModel.IndenterCommand = new IndentCommand(State, new Indenter(Vbe.Object, () => IndenterSettingsTests.GetMockIndenterSettings()), null);
            return this;
        }

        public MockedCodeExplorer ConfigureSaveDialog(string path, DialogResult result)
        {
            SaveDialog.Setup(o => o.FileName).Returns(path);
            SaveDialog.Setup(o => o.ShowDialog()).Returns(result);
            return this;
        }

        public MockedCodeExplorer ConfigureOpenDialog(string[] paths, DialogResult result)
        {
            OpenDialog.Setup(o => o.FileNames).Returns(paths);
            OpenDialog.Setup(o => o.ShowDialog()).Returns(result);
            return this;
        }

        public MockedCodeExplorer ConfigureFolderBrowser(string selected, DialogResult result)
        {
            FolderBrowser.Setup(m => m.SelectedPath).Returns(selected);
            FolderBrowser.Setup(m => m.ShowDialog()).Returns(result);
            return this;
        }

        public MockedCodeExplorer ConfigureMessageBox(ConfirmationOutcome result)
        {
            MessageBox.Setup(m => m.Confirm(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<ConfirmationOutcome>())).Returns(result);
            return this;
        }

        public MockedCodeExplorer SelectFirstProject()
        {
            ViewModel.SelectedItem = ViewModel.Projects.First();
            return this;
        }

        public MockedCodeExplorer SelectFirstCustomFolder()
        {
            ViewModel.SelectedItem = ViewModel.Projects.First().Children.First(node => node is CodeExplorerCustomFolderViewModel);
            return this;
        }

        public MockedCodeExplorer SelectFirstModule()
        {
            ViewModel.SelectedItem = ViewModel.Projects.First().Children.First(node => !(node is CodeExplorerReferenceFolderViewModel)).Children.First();
            return this;
        }

        public MockedCodeExplorer SelectFirstMember()
        {
            ViewModel.SelectedItem = ViewModel.Projects.First().Children.First(node => !(node is CodeExplorerReferenceFolderViewModel)).Children.First().Children.First();
            return this;
        }

        private Configuration GetDefaultUnitTestConfig()
        {
            var unitTestSettings = new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, true, true, false);

            var generalSettings = new GeneralSettings
            {
                EnableExperimentalFeatures = new List<ExperimentalFeature>
                    {
                        new ExperimentalFeature()
                    }
            };

            var userSettings = new UserSettings(generalSettings, null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                State?.Dispose();
            }
            _disposed = true;
        }
    }
}
