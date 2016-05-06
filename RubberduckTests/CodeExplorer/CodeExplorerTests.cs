using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.Commands;
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