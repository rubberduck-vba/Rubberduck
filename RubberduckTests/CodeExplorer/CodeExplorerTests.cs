using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
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

            var commands = new List<ICommand> {new CodeExplorer_AddStdModuleCommand(vbe.Object)};

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddStdModuleCommand.Execute(null);

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

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddClassModuleCommand.Execute(null);

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

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddUserFormCommand.Execute(null);

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

            var configLoader = new Mock<ConfigurationLoader>(null, null);
            configLoader.Setup(c => c.LoadConfiguration()).Returns(GetDefaultUnitTestConfig());

            var commands = new List<ICommand>
            {
                new CodeExplorer_AddTestModuleCommand(vbe.Object, new NewUnitTestModuleCommand(vbe.Object, configLoader.Object))
            };

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddTestModuleCommand.Execute(null);

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
            openFileDialog.Setup(o => o.FileName).Returns("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas");
            openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_ImportCommand(openFileDialog.Object)
            };

            var vm = new CodeExplorerViewModel(state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ImportCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Import("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
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

            var vm = new CodeExplorerViewModel(state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.ExportCommand.Execute(vm.SelectedItem);

            component.Verify(c => c.Export("C:\\Users\\Rubberduck\\Desktop\\StdModule1.bas"), Times.Once);
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

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.No);

            var commands = new List<ICommand>
            {
                new CodeExplorer_RemoveCommand(null, messageBox.Object)
            };

            var state = new RubberduckParserState();
            var vm = new CodeExplorerViewModel(state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.RemoveCommand.Execute(vm.SelectedItem);

            vbComponents.Verify(c => c.Remove(component), Times.Once);
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

            var vm = new CodeExplorerViewModel(state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First();
            vm.IndenterCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module.Lines());
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

            var state = new RubberduckParserState();
            var commands = new List<ICommand>
            {
                new CodeExplorer_RenameCommand(vbe.Object, state, new CodePaneWrapperFactory(), view.Object)
            };

            var vm = new CodeExplorerViewModel(state, commands);

            var parser = MockParser.Create(vbe.Object, state);
            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Projects.First().Items.First().Items.First().Items.First();
            vm.RenameCommand.Execute(vm.SelectedItem);

            Assert.AreEqual(expectedCode, module.Lines());
        }

        #region
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

            var userSettings = new UserSettings(null, null, null, unitTestSettings, null);
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
        #endregion
    }
}