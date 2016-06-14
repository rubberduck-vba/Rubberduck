using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
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
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.Extensions;
using RubberduckTests.Mocks;

namespace RubberduckTests.CodeExplorer
{
    [TestClass]
    public class CodeExplorerTests
    {
        [TestMethod]
        public void AddStdModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<ICommand> { new CodeExplorer_AddStdModuleCommand(vbe.Object) };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddStdModuleCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Add(vbext_ComponentType.vbext_ct_StdModule), Times.Once);
        }

        [TestMethod]
        public void AddClassModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<ICommand> { new CodeExplorer_AddClassModuleCommand(vbe.Object) };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddClassModuleCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Add(vbext_ComponentType.vbext_ct_ClassModule), Times.Once);
        }

        [TestMethod]
        public void AddUserForm()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<ICommand> { new CodeExplorer_AddUserFormCommand(vbe.Object) };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddUserFormCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Add(vbext_ComponentType.vbext_ct_MSForm), Times.Once);
        }

        [TestMethod]
        public void AddTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_AddTestModuleCommand(vbe.Object, new NewUnitTestModuleCommand(state, configLoader.Object))
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.AddTestModuleCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Add(vbext_ComponentType.vbext_ct_StdModule), Times.Once);
        }

        [TestMethod]
        public void ImportModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

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

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_ImportCommand(openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
        }

        [TestMethod]
        public void ImportMultipleModules()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

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

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_ImportCommand(openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
            vbComponents.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\ClsModule1.cls"), Times.Once);
        }

        [TestMethod]
        public void ImportModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = project.MockVBComponents;

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

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_ImportCommand(openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Never);
        }

        [TestMethod]
        public void ExportModule()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = projectMock.MockComponents.First();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_ExportCommand(saveFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ExportCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
        }

        [TestMethod]
        public void ExportModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = projectMock.MockComponents.First();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_ExportCommand(saveFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ExportCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Never);
        }

        [TestMethod]
        public void OpenDesigner()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectMock.AddComponent(projectMock.MockUserFormBuilder("UserForm1", "").Build());

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = projectMock.MockComponents.First();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_OpenDesignerCommand()
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.OpenDesignerCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.DesignerWindow(), Times.Once);
            Assert.IsTrue(component.Object.DesignerWindow().Visible);
        }

        [TestMethod]
        public void RemoveModule_Export()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>())).Returns(DialogResult.Yes);

            var commands = new List<ICommand>
            {
                new CodeExplorer_RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Remove(component), Times.Once);
        }

        [TestMethod]
        public void RemoveModule_Export_Cancel()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);
            saveFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            saveFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>())).Returns(DialogResult.Yes);

            var commands = new List<ICommand>
            {
                new CodeExplorer_RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Remove(component), Times.Never);
        }

        [TestMethod]
        public void RemoveModule_NoExport()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>())).Returns(DialogResult.No);

            var commands = new List<ICommand>
            {
                new CodeExplorer_RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Remove(component), Times.Once);
        }

        [TestMethod]
        public void RemoveModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, "");

            var vbComponents = projectMock.MockVBComponents;

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var saveFileDialog = new Mock<ISaveFileDialog>();
            saveFileDialog.Setup(o => o.OverwritePrompt);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>(), It.IsAny<MessageBoxDefaultButton>())).Returns(DialogResult.Cancel);

            var commands = new List<ICommand>
            {
                new CodeExplorer_RemoveCommand(saveFileDialog.Object, messageBox.Object)
            };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Remove(component), Times.Never);
        }

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
End Sub";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module.Lines());
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();

            Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
        }

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
End Sub";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component1 = project.Object.VBComponents.Item(0);
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents.Item(1);
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Lines());
            Assert.AreEqual(expectedCode, module2.Lines());
        }

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
End Sub";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode1)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component1 = project.Object.VBComponents.Item(0);
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents.Item(1);
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Lines());
            Assert.AreEqual(inputCode2, module2.Lines());
        }

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
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
        }

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
End Sub";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component1 = project.Object.VBComponents.Item(0);
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents.Item(1);
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Lines());
            Assert.AreEqual(expectedCode, module2.Lines());
        }

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
End Sub";

            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode1)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();
            var component1 = project.Object.VBComponents.Item(0);
            var module1 = component1.CodeModule;

            var component2 = project.Object.VBComponents.Item(1);
            var module2 = component2.CodeModule;

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module1.Lines());
            Assert.AreEqual(inputCode2, module2.Lines());
        }

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
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, inputCode);

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_IndentCommand(state, new Indenter(vbe.Object, GetDefaultIndenterSettings), null)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First();
            Assert.IsFalse(vm.IndenterCommand.CanExecute(vm.SelectedItem));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;

            var view = new Mock<IRenameDialog>();
            view.Setup(r => r.ShowDialog()).Returns(DialogResult.OK);
            view.Setup(r => r.Target);
            view.SetupGet(r => r.NewName).Returns("Fizz");

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_RenameCommand(vbe.Object, state, view.Object, msgbox.Object)
            };

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");
            vm.RenameCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void CompareByName_ReturnsZeroForIdenticalNodes()
        {
            var folderNode = new CodeExplorerCustomFolderViewModel(null, "Name", "Name");
            Assert.AreEqual(0, new CompareByName().Compare(folderNode, folderNode));
        }

        [TestMethod]
        public void CompareByName_ReturnsZeroForIdenticalNames()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name");
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name", "Name");

            Assert.AreEqual(0, new CompareByName().Compare(folderNode1, folderNode2));
        }

        [TestMethod]
        public void CompareByName_ReturnsCorrectOrdering()
        {
            // this won't happen, but just to be thorough...--besides, it is good for the coverage
            var folderNode1 = new CodeExplorerCustomFolderViewModel(null, "Name1", "Name1");
            var folderNode2 = new CodeExplorerCustomFolderViewModel(null, "Name2", "Name2");

            Assert.IsTrue(new CompareByName().Compare(folderNode1, folderNode2) < 0);
        }

        [TestMethod]
        public void CompareByType_ReturnsZeroForIdenticalNodes()
        {
            var errorNode = new CodeExplorerCustomFolderViewModel(null, "Name", "folder1.folder2");
            Assert.AreEqual(0, new CompareByName().Compare(errorNode, errorNode));
        }

        [TestMethod]
        public void CompareByType_ReturnsEventAboveConst()
        {
            var inputCode =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)
Public Const Bar = 0";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var eventNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
            var constNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar = 0");

            Assert.AreEqual(-1, new CompareByType().Compare(eventNode, constNode));
        }

        [TestMethod]
        public void CompareByType_ReturnsConstAboveField()
        {
            var inputCode =
@"Public Const Foo = 0
Public Bar As Boolean";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var constNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo = 0");
            var fieldNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareByType().Compare(constNode, fieldNode));
        }

        [TestMethod]
        public void CompareByType_ReturnsFieldAbovePropertyGet()
        {
            var inputCode =
@"Private Bar As Boolean

Public Property Get Foo() As Variant
End Property
";

            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var fieldNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");
            var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");

            Assert.AreEqual(-1, new CompareByType().Compare(fieldNode, propertyGetNode));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var propertyGetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Get)");
            var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");

            Assert.AreEqual(-1, new CompareByType().Compare(propertyGetNode, propertyLetNode));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var propertyLetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Let)");
            var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");

            Assert.AreEqual(-1, new CompareByType().Compare(propertyLetNode, propertySetNode));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var propertySetNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo (Set)");
            var functionNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareByType().Compare(propertySetNode, functionNode));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var functionNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Foo");
            var subNode = vm.Projects.First().Items.First().Items.First().Items.Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareByType().Compare(functionNode, subNode));
        }

        [TestMethod]
        public void CompareByType_ReturnsClassModuleBelowDocument()
        {
            var builder = new MockVbeBuilder();
            var projectMock = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("ClassModule1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .AddComponent("Sheet1", vbext_ComponentType.vbext_ct_Document, "");

            var project = projectMock.Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var docNode = vm.Projects.First().Items.First().Items.Single(s => s.Name == "Sheet1");
            var clsNode = vm.Projects.First().Items.First().Items.Single(s => s.Name == "ClassModule1");

            // this tests the logic I wrote to place docs above cls modules even though the parser calls them both cls modules
            Assert.AreEqual(((ICodeExplorerDeclarationViewModel) clsNode).Declaration.DeclarationType,
                ((ICodeExplorerDeclarationViewModel) docNode).Declaration.DeclarationType);

            Assert.AreEqual(-1, new CompareByType().Compare(docNode, clsNode));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");

            Assert.AreEqual(0, new CompareByName().Compare(vm.SelectedItem, vm.SelectedItem));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();

            var memberNode1 = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(s => s.Name == "Foo");
            var memberNode2 = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(s => s.Name == "Bar");

            Assert.AreEqual(-1, new CompareBySelection().Compare(memberNode1, memberNode2));
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var commands = new List<ICommand>();

            var vm = new CodeExplorerViewModel(new FolderHelper(state, GetDelimiterConfigLoader()), state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.OfType<CodeExplorerMemberViewModel>().Single(item => item.Declaration.IdentifierName == "Foo");

            Assert.AreEqual(0, new CompareByNodeType().Compare(vm.SelectedItem, vm.SelectedItem));
        }

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

            var userSettings = new UserSettings(null, null, null, null, unitTestSettings, null);
            return new Configuration(userSettings);
        }

        private Configuration GetDelimiterConfig()
        {
            var settings = new GeneralSettings
            {
                Delimiter = '.'
            };

            var userSettings = new UserSettings(settings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        private IIndenterSettings GetDefaultIndenterSettings()
        {
            var indenterSettings = new IndenterSettings
            {
                IndentEntireProcedureBody = true,
                IndentFirstCommentBlock = true,
                IndentFirstDeclarationBlock = true,
                AlignCommentsWithCode = true,
                AlignContinuations = true,
                IgnoreOperatorsInContinuations = true,
                IndentCase = false,
                ForceDebugStatementsInColumn1 = false,
                ForceCompilerDirectivesInColumn1 = false,
                IndentCompilerDirectives = true,
                AlignDims = false,
                AlignDimColumn = 15,
                EnableUndo = true,
                EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn,
                EndOfLineCommentColumnSpaceAlignment = 50,
                IndentSpaces = 4
            };

            return indenterSettings;
        }

        private ConfigurationLoader GetDelimiterConfigLoader()
        {
            var configLoader = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDelimiterConfig());

            return configLoader.Object;
        }
        #endregion
    }
}
