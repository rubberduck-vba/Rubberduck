using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEHost;
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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<ICommand> {new CodeExplorer_AddStdModuleCommand(vbe.Object)};

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddStdModuleCommand.Execute(null);

            Assert.IsTrue(vbe.Object.VBProjects.Item(0)
                    .VBComponents.Cast<VBComponent>()
                    .Count(c => c.Type == vbext_ComponentType.vbext_ct_StdModule) == 2);
        }

        [TestMethod]
        public void AddClassModule()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<ICommand> { new CodeExplorer_AddClassModuleCommand(vbe.Object) };

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddClassModuleCommand.Execute(null);

            var vbComponents = vbe.Object.VBProjects.Item(0).VBComponents.Cast<VBComponent>();

            Assert.IsTrue(vbComponents.Count(c => c.Type == vbext_ComponentType.vbext_ct_StdModule) == 1 &&
                vbComponents.Count(c => c.Type == vbext_ComponentType.vbext_ct_ClassModule) == 1);
        }

        [TestMethod]
        public void AddUserForm()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var commands = new List<ICommand> { new CodeExplorer_AddUserFormCommand(vbe.Object) };

            var vm = new CodeExplorerViewModel(new RubberduckParserState(), commands);
            vm.AddUserFormCommand.Execute(null);

            var vbComponents = vbe.Object.VBProjects.Item(0).VBComponents.Cast<VBComponent>();

            Assert.IsTrue(vbComponents.Count(c => c.Type == vbext_ComponentType.vbext_ct_StdModule) == 1 &&
                vbComponents.Count(c => c.Type == vbext_ComponentType.vbext_ct_MSForm) == 1);
        }

        [TestMethod]
        public void AddTestModule()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
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

            var vbComponents = vbe.Object.VBProjects.Item(0).VBComponents.Cast<VBComponent>();

            Assert.IsTrue(vbComponents.Count(c => c.Type == vbext_ComponentType.vbext_ct_StdModule) == 2);
        }

        [TestMethod]
        public void RemoveModule_Cancel()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .Build();
            var vbe = builder.AddProject(project).Build();

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

            var vbComponents = vbe.Object.VBProjects.Item(0).VBComponents.Cast<VBComponent>();
            Assert.IsTrue(vbComponents.Count(c => c.Type == vbext_ComponentType.vbext_ct_StdModule) == 0);
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
        #endregion
    }
}