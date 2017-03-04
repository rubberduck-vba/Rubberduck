using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.CodeExplorer
{
    [TestClass]
    public class CodeExplorerTests
    {
        [TestCategory("Code Explorer")]
        [TestMethod]
        public void AddStdModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<CommandBase> { new AddStdModuleCommand(vbe.Object) };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddStdModuleCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void AddClassModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<CommandBase> { new AddClassModuleCommand(vbe.Object) };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddClassModuleCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Add(ComponentType.ClassModule), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void AddUserForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<CommandBase> { new AddUserFormCommand(vbe.Object) };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddUserFormCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Add(ComponentType.UserForm), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void AddTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var state = new RubberduckParserState(vbe.Object);
            var vbeWrapper = vbe.Object;
            var commands = new List<CommandBase>
            {
                new Rubberduck.UI.CodeExplorer.Commands.AddTestModuleCommand(vbeWrapper, 
                    new Rubberduck.UI.Command.AddTestModuleCommand(vbeWrapper, state, configLoader.Object))
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddTestModuleCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Add(ComponentType.StandardModule), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ImportModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var openFileDialog = new Mock<IOpenFileDialog>();
            openFileDialog.Setup(o => o.AddExtension);
            openFileDialog.Setup(o => o.AutoUpgradeEnabled);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.Multiselect);
            openFileDialog.Setup(o => o.ShowHelp);
            openFileDialog.Setup(o => o.Filter);
            openFileDialog.Setup(o => o.CheckFileExists);
            openFileDialog.Setup(o => o.FileNames).Returns(new[] {"C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"});
            openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new ImportCommand(vbe.Object, openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ImportMultipleModules()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

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

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new ImportCommand(vbe.Object, openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
            components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\ClsModule1.cls"), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ImportModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

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

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new ImportCommand(vbe.Object, openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Never);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ExportModule()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = projectMock.MockComponents.First();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new ExportCommand(saveFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ExportCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ExportModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = projectMock.MockComponents.First();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new ExportCommand(saveFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ExportCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Never);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void OpenDesigner()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            projectMock.AddComponent(projectMock.MockUserFormBuilder("UserForm1", "").Build());

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = projectMock.MockComponents.First();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new OpenDesignerCommand()
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.OpenDesignerCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.DesignerWindow(), Times.Once);
            Assert.IsTrue(component.Object.DesignerWindow().IsVisible);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void RemoveCommand_RemovesModuleWhenPromptOk()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, string.Empty);

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[0];

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                                         It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>()))
                      .Returns(DialogResult.Yes);

            var commands = new List<CommandBase>
            {
                new RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Remove(component), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void RemoveCommand_CancelsWhenFilePromptCancels()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, string.Empty);

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[0];

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), 
                                         It.IsAny<MessageBoxButtons>(),
                                         It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>()))
                      .Returns(DialogResult.Yes);

            var commands = new List<CommandBase>
            {
                new RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Remove(component), Times.Never);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void RemoveCommand_GivenMsgBoxNO_RemovesModuleNoExport()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[0];

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), 
                                         It.IsAny<MessageBoxButtons>(),
                                         It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>()))
                      .Returns(DialogResult.No);

            var commands = new List<CommandBase>
            {
                new RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Remove(component), Times.Once);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void RemoveModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, "");

            var components = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[0];

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>())).Returns(DialogResult.Cancel);

            var commands = new List<CommandBase>
            {
                new RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            components.Verify(c => c.Remove(component), Times.Never);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object,() => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void IndentModule_DisabledWithNoIndentAnnotation()
        {
            var inputCode =
@"'@NoIndent

Sub Foo()
Dim d As Boolean
d = True
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

            Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Content());
            Assert.AreEqual(expectedCode, module2.Content());
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Content());
            Assert.AreEqual(inputCode2, module2.Content());
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Content());
            Assert.AreEqual(expectedCode, module2.Content());
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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
            var component1 = project.Object.VBComponents[0];
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents[1];
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Content());
            Assert.AreEqual(inputCode2, module2.Content());
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
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

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new IndentCommand(state, new Indenter(vbe.Object, () => Settings.IndenterSettingsTests.GetMockIndenterSettings()), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First();
            Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void RenameProcedure()
        {
            var inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            var expectedCode =
@"Sub Fizz()
End Sub

Sub Bar()
    Fizz
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;

            var view = new Mock<IRenameDialog>();
            view.Setup(r => r.ShowDialog()).Returns(DialogResult.OK);
            view.Setup(r => r.Target);
            view.SetupGet(r => r.NewName).Returns("Fizz");

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>
            {
                new RenameCommand(vbe.Object, state, view.Object, msgbox.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");
            vm.RenameCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ExpandAllNodes()
        {
            var inputCode =
@"Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, new List<CommandBase>());

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.Single();
            vm.ExpandAllSubnodesCommand.Execute(vm.SelectedItem);

            Assert.IsTrue(vm.Projects.Single().IsExpanded);
            Assert.IsTrue(vm.Projects.Single().Items.Single().IsExpanded);
            Assert.IsTrue(vm.Projects.Single().Items.Single().Items.Single().IsExpanded);
            Assert.IsTrue(vm.Projects.Single().Items.Single().Items.Single().Items.Single().IsExpanded);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void ExpandAllNodes_StartingWithSubnode()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Proj", ProjectProtection.Unprotected)
                .AddComponent("Comp1", ComponentType.ClassModule, @"'@Folder ""Foo""")
                .AddComponent("Comp2", ComponentType.ClassModule, @"'@Folder ""Bar""")
                .Build();
            var vbe = builder.AddProject(project).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, new List<CommandBase>());

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            
            vm.Projects.Single().Items.Last().IsExpanded = false;

            vm.SelectedItem = vm.Projects.Single().Items.First();
            vm.ExpandAllSubnodesCommand.Execute(vm.SelectedItem);
            
            Assert.IsTrue(vm.Projects.Single().Items.First().IsExpanded);
            Assert.IsFalse(vm.Projects.Single().Items.Last().IsExpanded);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CollapseAllNodes()
        {
            var inputCode =
@"Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, new List<CommandBase>());

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.Single();
            vm.CollapseAllSubnodesCommand.Execute(vm.SelectedItem);

            Assert.IsFalse(vm.Projects.Single().IsExpanded);
            Assert.IsFalse(vm.Projects.Single().Items.Single().IsExpanded);
            Assert.IsFalse(vm.Projects.Single().Items.Single().Items.Single().IsExpanded);
            Assert.IsFalse(vm.Projects.Single().Items.Single().Items.Single().Items.Single().IsExpanded);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CollapseAllNodes_StartingWithSubnode()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Proj", ProjectProtection.Unprotected)
                .AddComponent("Comp1", ComponentType.ClassModule, @"'@Folder ""Foo""")
                .AddComponent("Comp2", ComponentType.ClassModule, @"'@Folder ""Bar""")
                .Build();
            var vbe = builder.AddProject(project).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, new List<CommandBase>());

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.Projects.Single().Items.Last().IsExpanded = true;

            vm.SelectedItem = vm.Projects.Single().Items.First();
            vm.CollapseAllSubnodesCommand.Execute(vm.SelectedItem);

            Assert.IsFalse(vm.Projects.Single().Items.First().IsExpanded);
            Assert.IsTrue(vm.Projects.Single().Items.Last().IsExpanded);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByName_ReturnsZeroForIdenticalNodes()
        {
            var folderNode = new CodeExplorerCustomFolderViewModel(null, "Name", "Name");
            Assert.AreEqual(0, new CompareByName().Compare(folderNode, folderNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByName_ReturnsZeroForIdenticalNames()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name");
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name");

            Assert.AreEqual(0, new CompareByName().Compare(folderNode1, folderNode2));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByName_ReturnsCorrectOrdering()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name1", "Name1");
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name2", "Name2");

            Assert.IsTrue(new CompareByName().Compare(folderNode1, folderNode2) < 0);
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsZeroForIdenticalNodes()
        {
            var errorNode = new CodeExplorerCustomFolderViewModel(null, "Name", "folder1.folder2");
            Assert.AreEqual(0, new CompareByName().Compare(errorNode, errorNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsEventAboveConst()
        {
            var inputCode =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)
Public Const Bar = 0";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var eventNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
            var constNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar = 0");

            Assert.AreEqual(-1, new CompareByType().Compare(eventNode, constNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsConstAboveField()
        {
            var inputCode =
@"Public Const Foo = 0
Public Bar As Boolean";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var constNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo = 0");
            var fieldNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareByType().Compare(constNode, fieldNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsFieldAbovePropertyGet()
        {
            var inputCode =
@"Private Bar As Boolean

Public Property Get Foo() As Variant
End Property
";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var fieldNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");
            var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");

            Assert.AreEqual(-1, new CompareByType().Compare(fieldNode, propertyGetNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsPropertyGetAbovePropertyLet()
        {
            var inputCode =
@"Public Property Get Foo() As Variant
End Property

Public Property Let Foo(ByVal Value As Variant)
End Property
";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");
            var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");

            Assert.AreEqual(-1, new CompareByType().Compare(propertyGetNode, propertyLetNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsPropertyLetAbovePropertySet()
        {
            var inputCode =
@"Public Property Let Foo(ByVal Value As Variant)
End Property

Public Property Set Foo(ByVal Value As Variant)
End Property
";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");
            var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");

            Assert.AreEqual(-1, new CompareByType().Compare(propertyLetNode, propertySetNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsPropertySetAboveFunction()
        {
            var inputCode =
@"Public Property Set Foo(ByVal Value As Variant)
End Property

Public Function Bar() As Boolean
End Function
";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");
            var functionNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareByType().Compare(propertySetNode, functionNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsFunctionAboveSub()
        {
            var inputCode =
@"Public Function Foo() As Boolean
End Function

Public Sub Bar()
End Sub
";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var functionNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
            var subNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareByType().Compare(functionNode, subNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByType_ReturnsClassModuleBelowDocument()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("ClassModule1", ComponentType.ClassModule, "")
                .AddComponent("Sheet1", ComponentType.Document, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var docNode = vm.Projects.First().Items.First().Items.Single(s => s.Name == "Sheet1");
            var clsNode = vm.Projects.First().Items.First().Items.Single(s => s.Name == "ClassModule1");

            // this tests the logic I wrote to place docs above cls modules even though the parser calls them both cls modules
            Assert.AreEqual(((ICodeExplorerDeclarationViewModel) clsNode).Declaration.DeclarationType,
                ((ICodeExplorerDeclarationViewModel) docNode).Declaration.DeclarationType);

            Assert.AreEqual(-1, new CompareByType().Compare(docNode, clsNode));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareBySelection_ReturnsZeroForIdenticalNodes()
        {
            var inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");

            Assert.AreEqual(0, new CompareByName().Compare(vm.SelectedItem, vm.SelectedItem));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByNodeType_ReturnsCorrectMemberFirst_MemberPassedFirst()
        {
            var inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());

            var memberNode1 = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(s => s.Name == "Foo");
            var memberNode2 = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareBySelection().Compare(memberNode1, memberNode2));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByNodeType_ReturnsZeroForIdenticalNodes()
        {
            var inputCode =
@"Sub Foo()
End Sub

Sub Bar()
    Foo
End Sub";

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState(vbe.Object);
            var commands = new List<CommandBase>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");

            Assert.AreEqual(0, new CompareByNodeType().Compare(vm.SelectedItem, vm.SelectedItem));
        }

        [TestCategory("Code Explorer")]
        [TestMethod]
        public void CompareByNodeType_FoldersAreSortedByName()
        {
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "AAA", string.Empty);
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "zzz", string.Empty);

            Assert.IsTrue(new CompareByNodeType().Compare(folderNode1, folderNode2) < 0);
        }

        #region Helpers
        private Configuration GetDefaultUnitTestConfig()
        {
            var unitTestSettings = new UnitTestSettings
            {
                BindingMode = BindingMode.LateBinding,
                AssertMode = AssertMode.StrictAssert,
                ModuleInit = true,
                MethodInit = true,
                DefaultTestStubInNewModule = false
            };

            var userSettings = new UserSettings(null, null, null, null, unitTestSettings, null, null);
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
