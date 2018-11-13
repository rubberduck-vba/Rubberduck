using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UnitTesting;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class CodeExplorerTests
    {
        private GeneralSettings _generalSettings;
        private WindowSettings _windowSettings;

        private Mock<IConfigProvider<GeneralSettings>> _generalSettingsProvider;
        private Mock<IConfigProvider<WindowSettings>> _windowSettingsProvider;

        [SetUp]
        public void Initialize()
        {
            _generalSettings = new GeneralSettings();
            _windowSettings = new WindowSettings();

            _generalSettingsProvider = new Mock<IConfigProvider<GeneralSettings>>();
            _windowSettingsProvider = new Mock<IConfigProvider<WindowSettings>>();

            _generalSettingsProvider.Setup(s => s.Create()).Returns(_generalSettings);
            _windowSettingsProvider.Setup(s => s.Create()).Returns(_windowSettings);
        }

        [Category("Code Explorer")]
        [Test]
        public void AddStdModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();
            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var saveFileDialog = new Mock<ISaveFileDialog>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddStdModuleCommand = new AddStdModuleCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddStdModuleCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = true)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = true)]
        public bool AddStdModule_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");


            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddStdModuleCommand = new AddStdModuleCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddStdModuleCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddClassModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();            
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddClassModuleCommand = new AddClassModuleCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddClassModuleCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.ClassModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = true)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = true)]
        public bool AddClassModule_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");


            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);


            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddClassModuleCommand = new AddClassModuleCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddClassModuleCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddUserForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddUserFormCommand = new AddUserFormCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddUserFormCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.UserForm), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = false)]
        [TestCase(ProjectType.HostProject, ExpectedResult = true)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = true)]
        public bool AddUserForm_CanExecuteBasedOnProjectType(ProjectType projectType)
        {            
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");


            var vbe = builder.AddProject(project.Build()).Build();            
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {                
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddUserFormCommand = new AddUserFormCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddUserFormCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddVBForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddVBFormCommand = new AddVBFormCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddVBFormCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.VBForm), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddVBForm_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");


            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddVBFormCommand = new AddVBFormCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddVBFormCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddMDIForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddMDIFormCommand = new AddMDIFormCommand(vbe.Object, new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddMDIFormCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.MDIForm), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = false)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddMDIForm_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");


            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddMDIFormCommand = new AddMDIFormCommand(vbe.Object, new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddMDIFormCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddMDIForm_CannotExecuteIfProjectAlreadyHasMDIForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, ProjectType.StandardExe)
                .AddComponent("MDIForm1", ComponentType.MDIForm, "");


            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddMDIFormCommand = new AddMDIFormCommand(vbe.Object, new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                Assert.IsFalse(vm.AddMDIFormCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddUserControlForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddUserControlCommand = new AddUserControlCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddUserControlCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.UserControl), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddUserControl_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");


            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddUserControlCommand = new AddUserControlCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddUserControlCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddPropertyPage()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddPropertyPageCommand = new AddPropertyPageCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddPropertyPageCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.PropPage), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = true)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddPropertyPage_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddPropertyPageCommand = new AddPropertyPageCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddPropertyPageCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddUserDocument()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddUserDocumentCommand = new AddUserDocumentCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddUserDocumentCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.DocObject), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        [TestCase(ProjectType.StandardExe, ExpectedResult = false)]
        [TestCase(ProjectType.ActiveXExe, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXDll, ExpectedResult = true)]
        [TestCase(ProjectType.ActiveXControl, ExpectedResult = false)]
        [TestCase(ProjectType.HostProject, ExpectedResult = false)]
        [TestCase(ProjectType.StandAlone, ExpectedResult = false)]
        public bool AddUserDocument_CanExecuteBasedOnProjectType(ProjectType projectType)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected, projectType)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, null, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddUserDocumentCommand = new AddUserDocumentCommand(new AddComponentCommand(vbe.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                return vm.AddUserDocumentCommand.CanExecute(vm.SelectedItem);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            var projectRepository = new ProjectsRepository(vbe.Object);
            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vbeWrapper = vbe.Object;
                var uiDispatcher = new Mock<IUiDispatcher>();
                var interaction = new Mock<IVBEInteraction>();
                
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddTestModuleCommand = new AddTestModuleCommand(vbeWrapper, state, configLoader.Object, messageBox.Object, interaction.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddTestModuleCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vbeWrapper = vbe.Object;
                var uiDispatcher = new Mock<IUiDispatcher>();
                var interaction = new Mock<IVBEInteraction>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddTestModuleWithStubsCommand = new AddTestModuleWithStubsCommand(vbeWrapper, new AddTestModuleCommand(vbeWrapper, state, configLoader.Object, messageBox.Object, interaction.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.AddTestModuleWithStubsCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs_DisabledWhenParameterIsProject()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vbeWrapper = vbe.Object;
                var uiDispatcher = new Mock<IUiDispatcher>();
                var interaction = new Mock<IVBEInteraction>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddTestModuleWithStubsCommand = new AddTestModuleWithStubsCommand(vbeWrapper, new AddTestModuleCommand(vbeWrapper, state, configLoader.Object, messageBox.Object, interaction.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();

                Assert.IsFalse(vm.AddTestModuleWithStubsCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs_DisabledWhenParameterIsFolder()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vbeWrapper = vbe.Object;
                var interaction = new Mock<IVBEInteraction>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddTestModuleWithStubsCommand = new AddTestModuleWithStubsCommand(vbeWrapper, new AddTestModuleCommand(vbeWrapper, state, configLoader.Object, messageBox.Object, interaction.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First();

                Assert.IsFalse(vm.AddTestModuleWithStubsCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void AddTestModuleWithStubs_DisabledWhenParameterIsModuleMember()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "Public Sub S()\r\nEnd Sub");

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var vbeWrapper = vbe.Object;
                var uiDispatcher = new Mock<IUiDispatcher>();
                var interaction = new Mock<IVBEInteraction>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.AddTestModuleWithStubsCommand = new AddTestModuleWithStubsCommand(vbeWrapper, new AddTestModuleCommand(vbeWrapper, state, configLoader.Object, messageBox.Object, interaction.Object));

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.First();

                Assert.IsFalse(vm.AddTestModuleWithStubsCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var openFileDialog = new Mock<IOpenFileDialog>();
            openFileDialog.Setup(o => o.AddExtension);
            openFileDialog.Setup(o => o.AutoUpgradeEnabled);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.Multiselect);
            openFileDialog.Setup(o => o.ShowHelp);
            openFileDialog.Setup(o => o.Filter);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.FileNames).Returns(new[] { "C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas" });
            openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ImportCommand = new ImportCommand(vbe.Object, openFileDialog.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.ImportCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportMultipleModules()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var openFileDialog = new Mock<IOpenFileDialog>();
            openFileDialog.Setup(o => o.AddExtension);
            openFileDialog.Setup(o => o.AutoUpgradeEnabled);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.Multiselect);
            openFileDialog.Setup(o => o.ShowHelp);
            openFileDialog.Setup(o => o.Filter);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.FileNames).Returns(new[] { "C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas", "C:\\Users\\Rubberduck\\Desktop\\ClsModule1.cls" });
            openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ImportCommand = new ImportCommand(vbe.Object, openFileDialog.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.ImportCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
                components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\ClsModule1.cls"), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ImportModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var openFileDialog = new Mock<IOpenFileDialog>();
            openFileDialog.Setup(o => o.AddExtension);
            openFileDialog.Setup(o => o.AutoUpgradeEnabled);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.Multiselect);
            openFileDialog.Setup(o => o.ShowHelp);
            openFileDialog.Setup(o => o.Filter);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ImportCommand = new ImportCommand(vbe.Object, openFileDialog.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.ImportCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExportModule_ExpectExecution()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = projectMock.MockComponents.First();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ExportCommand = new ExportCommand(saveFileDialog.Object, state.ProjectsProvider);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.ExportCommand.Execute(vm.SelectedItem);

                component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExportModule_CancelPressed_ExpectNoExecution()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = projectMock.MockComponents.First();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ExportCommand = new ExportCommand(saveFileDialog.Object, state.ProjectsProvider);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.ExportCommand.Execute(vm.SelectedItem);

                component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Never);
            }
        }

        [Category("Commands")]
        [Test]
        public void ExportProject_TestCanExecute_ExpectTrue()
        {
            var builder = new MockVbeBuilder();

            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var project = projectMock.Build();

            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var mockFolderBrowser = new Mock<IFolderBrowser>();
            var mockFolderBrowserFactory = new Mock<IFolderBrowserFactory>();

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ExportAllCommand = new ExportAllCommand(vbe.Object, mockFolderBrowserFactory.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                project.SetupGet(m => m.VBComponents.Count).Returns(1);
                vm.SelectedItem = vm.Projects.First();

                Assert.IsTrue(vm.ExportAllCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Commands")]
        [Test]
        public void ExportProject_TestExecute_OKPressed_ExpectExecution()
        {
            string path = @"C:\Users\Rubberduck\Desktop\ExportAll";
            string projectPath = @"C:\Users\Rubberduck\Documents\Subfolder";
            string projectFullPath = @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm";

            var builder = new MockVbeBuilder();

            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var project = projectMock.Build();
            project.SetupGet(m => m.IsSaved).Returns(true);
            project.SetupGet(m => m.FileName).Returns(projectFullPath);

            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;
            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var mockFolderBrowser = new Mock<IFolderBrowser>();
            mockFolderBrowser.Setup(m => m.SelectedPath).Returns(path);
            mockFolderBrowser.Setup(m => m.ShowDialog()).Returns(DialogResult.OK);

            var mockFolderBrowserFactory = new Mock<IFolderBrowserFactory>();
            mockFolderBrowserFactory.Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true, projectPath)).Returns(mockFolderBrowser.Object);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ExportAllCommand = new ExportAllCommand(null, mockFolderBrowserFactory.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.ExportAllCommand.Execute(vm.SelectedItem);

                project.Verify(m => m.ExportSourceFiles(path), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void ExportProject_TestExecute_CancelPressed_ExpectExecution()
        {
            string path = @"C:\Users\Rubberduck\Desktop\ExportAll";
            string projectPath = @"C:\Users\Rubberduck\Documents\Subfolder";
            string projectFullPath = @"C:\Users\Rubberduck\Documents\Subfolder\Project.xlsm";

            var builder = new MockVbeBuilder();

            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "")
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Document1", ComponentType.Document, "")
                .AddComponent("UserForm1", ComponentType.UserForm, "");

            var project = projectMock.Build();
            project.SetupGet(m => m.IsSaved).Returns(true);
            project.SetupGet(m => m.FileName).Returns(projectFullPath);

            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;
            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var mockFolderBrowser = new Mock<IFolderBrowser>();
            mockFolderBrowser.Setup(m => m.SelectedPath).Returns(path);
            mockFolderBrowser.Setup(m => m.ShowDialog()).Returns(DialogResult.Cancel);

            var mockFolderBrowserFactory = new Mock<IFolderBrowserFactory>();
            mockFolderBrowserFactory.Setup(m => m.CreateFolderBrowser(It.IsAny<string>(), true, projectPath)).Returns(mockFolderBrowser.Object);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.ExportAllCommand = new ExportAllCommand(null, mockFolderBrowserFactory.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.ExportAllCommand.Execute(vm.SelectedItem);

                project.Verify(m => m.ExportSourceFiles(path), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void OpenDesigner()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            projectMock.MockUserFormBuilder("UserForm1", "").AddFormToProjectBuilder();

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = projectMock.MockComponents.First();

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.OpenDesignerCommand = new OpenDesignerCommand(state.ProjectsProvider);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.OpenDesignerCommand.Execute(vm.SelectedItem);

                component.Verify(c => c.DesignerWindow(), Times.Once);
                Assert.IsTrue(component.Object.DesignerWindow().IsVisible);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveCommand_RemovesModuleWhenPromptOk()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, string.Empty);

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = project.Object.VBComponents[0];


            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Confirm(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<ConfirmationOutcome>())).Returns(ConfirmationOutcome.Yes);

            var projectRepository = new ProjectsRepository(vbe.Object);
            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.RemoveCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Remove(component), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveCommand_CancelsWhenFilePromptCancels()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, string.Empty);

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = project.Object.VBComponents[0];
            
            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(true);

            var projectRepository = new ProjectsRepository(vbe.Object);
            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var commands = new List<CommandBase>
                {
                    new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider)
                };
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.RemoveCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Remove(component), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveCommand_GivenMsgBoxNO_RemovesModuleNoExport()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = project.Object.VBComponents[0];

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Confirm(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<ConfirmationOutcome>())).Returns(ConfirmationOutcome.No);

            var projectRepository = new ProjectsRepository(vbe.Object);
            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.RemoveCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Remove(component), Times.Once);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void RemoveModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component = project.Object.VBComponents[0];

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Confirm(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<ConfirmationOutcome>())).Returns(ConfirmationOutcome.Cancel);

            var projectRepository = new ProjectsRepository(vbe.Object);
            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var commands = new List<CommandBase>
                {
                    new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider)
                };
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.RemoveCommand.Execute(vm.SelectedItem);

                components.Verify(c => c.Remove(component), Times.Never);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentModule()
        {
            var inputCode =
                @"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.IndenterCommand.Execute(vm.SelectedItem);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentModule_DisabledWithNoIndentAnnotation()
        {
            var inputCode =
                @"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

                Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentProject()
        {
            var inputCode =
                @"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddComponent("ClassModule1", ComponentType.ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.IndenterCommand.Execute(vm.SelectedItem);

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(expectedCode, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentProject_IndentsModulesWithoutNoIndentAnnotation()
        {
            var inputCode1 =
                @"Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var inputCode2 =
                @"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode1)
                .AddComponent("ClassModule1", ComponentType.ClassModule, inputCode2);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.IndenterCommand.Execute(vm.SelectedItem);

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(inputCode2, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentProject_DisabledWhenAllModulesHaveNoIndentAnnotation()
        {
            var inputCode =
                @"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddComponent("ClassModule1", ComponentType.ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First();
                vm.IndenterCommand.Execute(vm.SelectedItem);

                Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentFolder()
        {
            var inputCode =
                @"'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var expectedCode =
                @"'@Folder ""folder""

Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddComponent("ClassModule1", ComponentType.ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First();
                vm.IndenterCommand.Execute(vm.SelectedItem);

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(expectedCode, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentFolder_IndentsModulesWithoutNoIndentAnnotation()
        {
            var inputCode1 =
                @"'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var inputCode2 =
                @"'@NoIndent
'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var expectedCode =
                @"'@Folder ""folder""

Sub Foo()
    Dim d As Boolean
    d = True
End Sub
";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode1)
                .AddComponent("ClassModule1", ComponentType.ClassModule, inputCode2);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First();
                vm.IndenterCommand.Execute(vm.SelectedItem);

                Assert.AreEqual(expectedCode, module1.Content());
                Assert.AreEqual(inputCode2, module2.Content());
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void IndentFolder_DisabledWhenAllModulesHaveNoIndentAnnotation()
        {
            var inputCode =
                @"'@NoIndent
'@Folder ""folder""

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddComponent("ClassModule1", ComponentType.ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);
                vm.IndenterCommand = new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First();
                Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExpandAllNodes()
        {
            var inputCode =
                @"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.Single();
                vm.ExpandAllSubnodesCommand.Execute(vm.SelectedItem);

                Assert.IsTrue(vm.Projects.Single().IsExpanded);
                Assert.IsTrue(vm.Projects.Single().Items.Single().IsExpanded);
                Assert.IsTrue(vm.Projects.Single().Items.Single().Items.Single().IsExpanded);
                Assert.IsTrue(vm.Projects.Single().Items.Single().Items.Single().Items.Single().IsExpanded);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void ExpandAllNodes_StartingWithSubnode()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Proj", ProjectProtection.Unprotected)
                .AddComponent("Comp1", ComponentType.ClassModule, @"'@Folder ""Foo""")
                .AddComponent("Comp2", ComponentType.ClassModule, @"'@Folder ""Bar""")
                .Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.Projects.Single().Items.Last().IsExpanded = false;

                vm.SelectedItem = vm.Projects.Single().Items.First();
                vm.ExpandAllSubnodesCommand.Execute(vm.SelectedItem);

                Assert.IsTrue(vm.Projects.Single().Items.First().IsExpanded);
                Assert.IsFalse(vm.Projects.Single().Items.Last().IsExpanded);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CollapseAllNodes()
        {
            var inputCode =
                @"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.Single();
                vm.CollapseAllSubnodesCommand.Execute(vm.SelectedItem);

                Assert.IsFalse(vm.Projects.Single().IsExpanded);
                Assert.IsFalse(vm.Projects.Single().Items.Single().IsExpanded);
                Assert.IsFalse(vm.Projects.Single().Items.Single().Items.Single().IsExpanded);
                Assert.IsFalse(vm.Projects.Single().Items.Single().Items.Single().Items.Single().IsExpanded);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CollapseAllNodes_StartingWithSubnode()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Proj", ProjectProtection.Unprotected)
                .AddComponent("Comp1", ComponentType.ClassModule, @"'@Folder ""Foo""")
                .AddComponent("Comp2", ComponentType.ClassModule, @"'@Folder ""Bar""")
                .Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.Projects.Single().Items.Last().IsExpanded = true;

                vm.SelectedItem = vm.Projects.Single().Items.First();
                vm.CollapseAllSubnodesCommand.Execute(vm.SelectedItem);

                Assert.IsFalse(vm.Projects.Single().Items.First().IsExpanded);
                Assert.IsTrue(vm.Projects.Single().Items.Last().IsExpanded);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByName_NotAlreadySelectedInMenu_ExpectTrue()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = false,
                    CodeExplorer_SortByCodeOrder = true
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetNameSortCommand.Execute(true);

                Assert.IsTrue(vm.SortByName);
                Assert.IsFalse(vm.SortByCodeOrder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByName_AlreadySelectedInMenu_ExpectTrue()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = true,
                    CodeExplorer_SortByCodeOrder = false
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetNameSortCommand.Execute(false);

                Assert.IsTrue(vm.SortByName);
                Assert.IsFalse(vm.SortByCodeOrder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByName_BothSortOptionsFalse_ExpectTrueOnlyForSortByName()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = false,
                    CodeExplorer_SortByCodeOrder = false
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetNameSortCommand.Execute(true);

                Assert.IsTrue(vm.SortByName);
                Assert.IsFalse(vm.SortByCodeOrder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByName_BothSortOptionsTrue_ExpectTrueOnlyForSortByName()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var uiDispatcher = new Mock<IUiDispatcher>();

            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = true,
                    CodeExplorer_SortByCodeOrder = true
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetNameSortCommand.Execute(true);

                Assert.IsTrue(vm.SortByName);
                Assert.IsFalse(vm.SortByCodeOrder);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByCodeOrder_NotAlreadySelectedInMenu_ExpectTrue()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = true,
                    CodeExplorer_SortByCodeOrder = false
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetCodeOrderSortCommand.Execute(true);

                Assert.IsTrue(vm.SortByCodeOrder);
                Assert.IsFalse(vm.SortByName);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByCodeOrder_AlreadySelectedInMenu_ExpectTrue()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = false,
                    CodeExplorer_SortByCodeOrder = true
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var uiDispatcher = new Mock<IUiDispatcher>();
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetCodeOrderSortCommand.Execute(false);

                Assert.IsTrue(vm.SortByCodeOrder);
                Assert.IsFalse(vm.SortByName);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByCodeOrder_BothSortOptionsFalse_ExpectCorrectSortPair()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = false,
                    CodeExplorer_SortByCodeOrder = false
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetCodeOrderSortCommand.Execute(true);

                Assert.IsTrue(vm.SortByCodeOrder);
                Assert.IsFalse(vm.SortByName);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void SetSortByCodeOrder_BothSortOptionsTrue_ExpectCorrectSortPair()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var commands = new List<CommandBase> { new AddStdModuleCommand(new AddComponentCommand(vbe.Object)) };

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {

                var windowSettings = new WindowSettings
                {
                    CodeExplorer_SortByName = true,
                    CodeExplorer_SortByCodeOrder = true
                };
                _windowSettingsProvider.Setup(s => s.Create()).Returns(windowSettings);

                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
                vm.SetCodeOrderSortCommand.Execute(true);

                Assert.IsTrue(vm.SortByCodeOrder);
                Assert.IsFalse(vm.SortByName);
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsZeroForIdenticalNodes()
        {
            var folderNode = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, null);
            Assert.AreEqual(0, new CompareByName().Compare(folderNode, folderNode));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsZeroForIdenticalNames()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, null);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name", null, null);

            Assert.AreEqual(0, new CompareByName().Compare(folderNode1, folderNode2));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByName_ReturnsCorrectOrdering()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name1", "Name1", null, null);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name2", "Name2", null, null);

            Assert.IsTrue(new CompareByName().Compare(folderNode1, folderNode2) < 0);
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsZeroForIdenticalNodes()
        {
            var errorNode = new CodeExplorerCustomFolderViewModel(null, "Name", "folder1.folder2", null, null);
            Assert.AreEqual(0, new CompareByName().Compare(errorNode, errorNode));
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsEventAboveConst()
        {
            var inputCode =
                @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)
Public Const Bar = 0";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var eventNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
                var constNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar = 0");

                Assert.AreEqual(-1, new CompareByType().Compare(eventNode, constNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsConstAboveField()
        {
            var inputCode =
                @"Public Const Foo = 0
Public Bar As Boolean";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var constNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo = 0");
                var fieldNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByType().Compare(constNode, fieldNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsFieldAbovePropertyGet()
        {
            var inputCode =
                @"Private Bar As Boolean

Public Property Get Foo() As Variant
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var fieldNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");
                var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");

                Assert.AreEqual(-1, new CompareByType().Compare(fieldNode, propertyGetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyGetEqualToPropertyLet()
        {
            var inputCode =
                @"Public Property Get Foo() As Variant
End Property

Public Property Let Foo(ByVal Value As Variant)
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");
                var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyGetNode, propertyLetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyGetEqualToPropertySet()
        {
            var inputCode =
                @"Public Property Get Foo() As Variant
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");
                var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyGetNode, propertyLetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyLetEqualToPropertyGet()
        {
            var inputCode =
                @"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Get Foo() As Variant
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");
                var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyLetNode, propertySetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertyLetEqualToPropertySet()
        {
            var inputCode =
                @"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");
                var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");

                Assert.AreEqual(0, new CompareByType().Compare(propertyLetNode, propertySetNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPropertySetAboveFunction()
        {
            var inputCode =
                @"Public Property Set Foo(ByVal Value As Variant)
End Property

Public Function Bar() As Boolean
End Function
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");
                var functionNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByType().Compare(propertySetNode, functionNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsSubsAndFunctionsEqual()
        {
            var inputCode =
                @"Public Function Foo() As Boolean
End Function

Public Sub Bar()
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var functionNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
                var subNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(0, new CompareByType().Compare(functionNode, subNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsPublicMethodsAbovePrivateMethods()
        {
            var inputCode =
                @"Private Sub Foo()
End Sub

Public Sub Bar()
End Sub
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var privateNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
                var publicNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareByType().Compare(publicNode, privateNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByType_ReturnsClassModuleBelowDocument()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Sheet1", ComponentType.Document, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var docNode = vm.Projects.First().Items.First().Items.Single(s => s.Name == "Sheet1");
                var clsNode = vm.Projects.First().Items.First().Items.Single(s => s.Name == "ClassModule1");

                // this tests the logic I wrote to place docs above cls modules even though the parser calls them both cls modules
                Assert.AreEqual(((ICodeExplorerDeclarationViewModel)clsNode).Declaration.DeclarationType,
                    ((ICodeExplorerDeclarationViewModel)docNode).Declaration.DeclarationType);

                Assert.AreEqual(-1, new CompareByType().Compare(docNode, clsNode));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareBySelection_ReturnsZeroForIdenticalNodes()
        {
            var inputCode =
                @"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");

                Assert.AreEqual(0, new CompareByName().Compare(vm.SelectedItem, vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_ReturnsCorrectMemberFirst_MemberPassedFirst()
        {
            var inputCode =
                @"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());

                var memberNode1 = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(s => s.Name == "Foo");
                var memberNode2 = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(s => s.Name == "Bar");

                Assert.AreEqual(-1, new CompareBySelection().Compare(memberNode1, memberNode2));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_ReturnsZeroForIdenticalNodes()
        {
            var inputCode =
                @"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var projectRepository = new ProjectsRepository(vbe.Object);
            var messageBox = new Mock<IMessageBox>();
            var saveFileDialog = new Mock<ISaveFileDialog>();

            using (var state = new RubberduckParserState(vbe.Object, projectRepository, new DeclarationFinderFactory(), vbeEvents.Object))
            {
                var uiDispatcher = new Mock<IUiDispatcher>();
                var removeCommand = new RemoveCommand(saveFileDialog.Object, messageBox.Object, state.ProjectsProvider);
                var vm = new CodeExplorerViewModel(new FolderHelper(state, vbe.Object), state, removeCommand, _generalSettingsProvider.Object, _windowSettingsProvider.Object, uiDispatcher.Object, vbe.Object);

                var parser = MockParser.Create(vbe.Object, state, projectRepository);
                parser.Parse(new CancellationTokenSource());
                if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

                vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");

                Assert.AreEqual(0, new CompareByNodeType().Compare(vm.SelectedItem, vm.SelectedItem));
            }
        }

        [Category("Code Explorer")]
        [Test]
        public void CompareByNodeType_FoldersAreSortedByName()
        {
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "AAA", string.Empty, null, null);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "zzz", string.Empty, null, null);

            Assert.IsTrue(new CompareByNodeType().Compare(folderNode1, folderNode2) < 0);
        }

        #region Helpers
        private Configuration GetDefaultUnitTestConfig()
        {
            var unitTestSettings = new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, true, true, false);

            var generalSettings = new GeneralSettings
            {
                EnableExperimentalFeatures = new List<ExperimentalFeatures>
                {
                    new ExperimentalFeatures()
                }
            };

            var userSettings = new UserSettings(generalSettings, null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }

        //private Configuration GetDelimiterConfig()
        //{
        //    var settings = new GeneralSettings
        //    {
        //        Delimiter = '.'
        //    };

        //    var userSettings = new UserSettings(settings, null, null, null, null, null, null);
        //    return new Configuration(userSettings);
        //}

        //private ConfigurationLoader GetDelimiterConfigLoader()
        //{
        //    var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null);
        //    configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDelimiterConfig());

        //    return configLoader.Object;
        //}
        #endregion
    }
}
