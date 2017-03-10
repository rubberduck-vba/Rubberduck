using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class UnitTestCommandTests
    {
        [TestCategory("Commands")]
        [TestMethod]
        public void AddsTest()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule
Private Assert As Object
{0}";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

            addTestMethodCommand.Execute(null);
            var module = component.CodeModule;

            Assert.AreEqual(
                string.Format(input,
                    AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder, "TestMethod1")) +
                Environment.NewLine, module.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddsTest_NullActiveCodePane()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule
Private Assert As Object
";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

            addTestMethodCommand.Execute(null);
            var module = component.CodeModule;

            Assert.AreEqual(input, module.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddTest_CanExecute_NonReadyState()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences);

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
            Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddTest_CanExecute_NoTestModule()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
            Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddTest_CanExecute()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule
Private Assert As Object
";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
            Assert.IsTrue(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddsExpectedErrorTest()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule
Private Assert As Object
{0}";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

            addTestMethodCommand.Execute(null);
            var module = component.CodeModule;

            Assert.AreEqual(
                string.Format(input,
                    AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
                        "TestMethod1")) + Environment.NewLine, module.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddExpectedErrorTest_CanExecute_NonReadyState()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences);

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
            Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddExpectedErrorTest_CanExecute_NoTestModule()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
            Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddExpectedErrorTest_CanExecute()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule
Private Assert As Object
";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
            Assert.IsTrue(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddsExpectedErrorTest_NullActiveCodePane()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule
Private Assert As Object
";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
            addTestMethodCommand.Execute(null);
            
            Assert.AreEqual(input, component.CodeModule.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddsTestModule()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null);
            var config = GetUnitTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object);
            addTestModuleCommand.Execute(null);

            // mock suite auto-assigns "TestModule1" to the first component when we create the mock
            var project = state.DeclarationFinder.FindProject("TestProject1");
            var module = state.DeclarationFinder.FindStdModule("TestModule2", project);
            Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        }

        private Configuration GetUnitTestConfig()
        {
            var unitTestSettings = new UnitTestSettings
            {
                BindingMode = BindingMode.EarlyBinding,
                AssertMode = AssertMode.StrictAssert,
                ModuleInit = false,
                MethodInit = false,
                DefaultTestStubInNewModule = false
            };

            var userSettings = new UserSettings(null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }
    }
}