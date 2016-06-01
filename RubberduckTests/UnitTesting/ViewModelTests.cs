using System.Linq;
using System.Threading;
using System.Windows.Media;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    [TestClass]
    public class ViewModelTests
    {
        [TestMethod]
        public void UIDiscoversAnnotatedTestMethods()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            UiDispatcher.Initialize();

            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var model = new TestExplorerModel(vbe.Object, parser.State);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(1, model.Tests.Count);
        }

        [Ignore]    // it fails sporadically when not run by itself--figure out what's up.
        [TestMethod]
        public void UIRemovesRemovedTestMethods()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            UiDispatcher.Initialize();

            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods)
                .AddComponent("TestModule2", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var model = new TestExplorerModel(vbe.Object, parser.State);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            
            Assert.AreEqual(2, model.Tests.Count);

            project.MockComponents.RemoveAt(1);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(1, model.Tests.Count);
        }

        [TestMethod]
        public void UISetsProgressBarColor_GreenForSuccess()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            UiDispatcher.Initialize();

            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var model = new TestExplorerModel(vbe.Object, parser.State);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(1, model.Tests.Count);

            model.Tests.First().Result = new TestResult(TestOutcome.Succeeded);
            model.AddExecutedTest(model.Tests.First());

            Assert.AreEqual(model.ProgressBarColor, Colors.LimeGreen);
        }

        [TestMethod]
        public void UISetsProgressBarColor_RedForFailure()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            UiDispatcher.Initialize();

            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var model = new TestExplorerModel(vbe.Object, parser.State);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(1, model.Tests.Count);

            model.Tests.First().Result = new TestResult(TestOutcome.Failed);
            model.AddExecutedTest(model.Tests.First());

            Assert.AreEqual(model.ProgressBarColor, Colors.Red);
        }

        [TestMethod]
        public void UISetsProgressBarColor_YellowForInconclusive()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            UiDispatcher.Initialize();

            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var model = new TestExplorerModel(vbe.Object, parser.State);

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(1, model.Tests.Count);

            model.Tests.First().Result = new TestResult(TestOutcome.Inconclusive);
            model.AddExecutedTest(model.Tests.First());

            Assert.AreEqual(model.ProgressBarColor, Colors.Gold);
        }

        private const string RawInput = @"Option Explicit
Option Private Module

{0}
Private Assert As New Rubberduck.AssertClass

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
";

        private string GetTestModuleInput
        {
            get { return string.Format(RawInput, "'@TestModule"); }
        }
    }
}