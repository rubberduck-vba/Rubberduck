using System;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, parser.State);

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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, parser.State);

            addTestMethodCommand.Execute(null);
            var module = component.CodeModule;

            Assert.AreEqual(input, module.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddTest_CanExecute_NonReadyState()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            parser.State.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences);

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, parser.State);
            Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddTest_CanExecute_NoTestModule()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, parser.State);
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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, parser.State);

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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, parser.State);

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
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            parser.State.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences);

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, parser.State);
            Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddExpectedErrorTest_CanExecute_NoTestModule()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, parser.State);
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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, parser.State);
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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, parser.State);
            addTestMethodCommand.Execute(null);

            var module = component.CodeModule;
            Assert.AreEqual(input, module.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddsTestModule()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null);
            var config = GetUnitTestConfig();
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, parser.State, settings.Object);
            addTestModuleCommand.Execute(null);

            // mock suite auto-assigns "TestModule1" to the first component when we create the mock
            var project = parser.State.DeclarationFinder.FindProject("TestProject1");
            var module = parser.State.DeclarationFinder.FindStdModule("TestModule2", project);
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