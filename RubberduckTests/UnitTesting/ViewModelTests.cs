using System.Linq;
using System.Threading;
using System.Windows.Media;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    [TestClass]
    public class ViewModelTests
    {
        [TestMethod]
        public void UIDiscoversAnnotatedTestMethods()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder()
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods)
                .MockVbeBuilder();

            var vbe = builder.Build().Object;

            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));
            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            
            Assert.AreEqual(1, model.Tests.Count);
        }

        [TestMethod]
        public void UIRemovesRemovedTestMethods()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput + testMethods);
            builder.AddProject(project.Build());

            var vbe = builder.Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            
            Assert.AreEqual(2, model.Tests.Count);

            project.RemoveComponent(project.MockComponents[1]);
            
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            
            Assert.AreEqual(1, model.Tests.Count);
        }

        [TestMethod]
        public void UISetsProgressBarColor_LimeGreenForSuccess()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder
                .ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests.First().Result = new TestResult(TestOutcome.Succeeded);
            model.AddExecutedTest(model.Tests.First());

            Assert.AreEqual(model.ProgressBarColor, Colors.LimeGreen);
        }

        [TestMethod]
        public void UISetsProgressBarColor_RedForFailure()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests.First().Result = new TestResult(TestOutcome.Failed);
            model.AddExecutedTest(model.Tests.First());

            Assert.AreEqual(model.ProgressBarColor, Colors.Red);
        }

        [TestMethod]
        public void UISetsProgressBarColor_GoldForInconclusive()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests.First().Result = new TestResult(TestOutcome.Inconclusive);
            model.AddExecutedTest(model.Tests.First());

            Assert.AreEqual(model.ProgressBarColor, Colors.Gold);
        }

        [TestMethod]
        public void UISetsProgressBarColor_RedForFailure_IncludesNonFailingTests()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub

'@TestMethod
Public Sub TestMethod2()
End Sub

'@TestMethod
Public Sub TestMethod3()
End Sub

'@TestMethod
Public Sub TestMethod4()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests[0].Result = new TestResult(TestOutcome.Succeeded);
            model.Tests[1].Result = new TestResult(TestOutcome.Inconclusive);
            model.Tests[2].Result = new TestResult(TestOutcome.Failed);
            model.Tests[3].Result = new TestResult(TestOutcome.Ignored);

            model.AddExecutedTest(model.Tests[0]);
            model.AddExecutedTest(model.Tests[1]);
            model.AddExecutedTest(model.Tests[2]);
            model.AddExecutedTest(model.Tests[3]);

            Assert.AreEqual(model.ProgressBarColor, Colors.Red);
        }

        [TestMethod]
        public void UISetsProgressBarColor_GoldForInconclusive_IncludesNonFailingAndNonInconclusiveTests()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub

'@TestMethod
Public Sub TestMethod2()
End Sub

'@TestMethod
Public Sub TestMethod3()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests[0].Result = new TestResult(TestOutcome.Succeeded);
            model.Tests[1].Result = new TestResult(TestOutcome.Inconclusive);
            model.Tests[2].Result = new TestResult(TestOutcome.Ignored);

            model.AddExecutedTest(model.Tests[0]);
            model.AddExecutedTest(model.Tests[1]);
            model.AddExecutedTest(model.Tests[2]);

            Assert.AreEqual(model.ProgressBarColor, Colors.Gold);
        }

        [TestMethod]
        public void UISetsProgressBarColor_LimeGreenForSuccess_IncludesIgnoredTests()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub

'@TestMethod
Public Sub TestMethod2()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests[0].Result = new TestResult(TestOutcome.Succeeded);
            model.Tests[1].Result = new TestResult(TestOutcome.Ignored);

            model.AddExecutedTest(model.Tests[0]);
            model.AddExecutedTest(model.Tests[1]);

            Assert.AreEqual(model.ProgressBarColor, Colors.LimeGreen);
        }

        [TestMethod]
        public void AddingExecutedTestUpdatesExecutedCount()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(0, model.ExecutedCount);

            model.Tests[0].Result = new TestResult(TestOutcome.Succeeded);
            model.AddExecutedTest(model.Tests[0]);

            Assert.AreEqual(1, model.ExecutedCount);
        }

        [TestMethod]
        public void AddingExecutedTestUpdatesLastRun()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(0, model.LastRun.Count);

            model.Tests[0].Result = new TestResult(TestOutcome.Succeeded);
            model.AddExecutedTest(model.Tests[0]);

            Assert.AreEqual(1, model.LastRun.Count);
        }

        [TestMethod]
        public void ClearLastRun()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            var model = new TestExplorerModel(vbe, parser.State);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            model.Tests[0].Result = new TestResult(TestOutcome.Succeeded);
            model.AddExecutedTest(model.Tests[0]);

            Assert.AreEqual(1, model.LastRun.Count);

            model.ClearLastRun();

            Assert.AreEqual(0, model.LastRun.Count);
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