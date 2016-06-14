using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    [TestClass]
    public class DiscoveryTests
    {
        [TestMethod]
        public void Discovery_DiscoversAnnotatedTestMethods()
        {
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

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.AreEqual(1, UnitTestHelpers.GetAllTests(vbe.Object, parser.State).Count());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedTestMethods()
        {
            var testMethods = @"Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.IsFalse(UnitTestHelpers.GetAllTests(vbe.Object, parser.State).Any());
        }

        [TestMethod]
        public void Discovery_IgnoresAnnotatedTestMethodsNotInTestModule()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetNormalModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            Assert.IsFalse(UnitTestHelpers.GetAllTests(vbe.Object, parser.State).Any());
        }

        [TestMethod]
        public void Discovery_DiscoversAnnotatedTestMethodsInGivenTestModule()
        {
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

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var tests = project.MockComponents.Single(f => f.Object.Name == "TestModule1")
                    .Object.GetTests(vbe.Object, parser.State)
                    .ToList();

            Assert.AreEqual(1, tests.Count);
            Assert.AreEqual("TestModule1", tests.ElementAt(0).Declaration.ComponentName);
        }

        [TestMethod]
        public void Discovery_DiscoversAnnotatedTestInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput)
                .AddComponent("TestModule2", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindTestInitializeMethods(parser.State).ToList();

            Assert.AreEqual(1, initMethods.Count);
            Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedModuleName.ComponentName);
            Assert.AreEqual("TestInitialize", initMethods.ElementAt(0).MemberName);
        }

        [TestMethod]
        public void Discovery_DiscoversAnnotatedTestCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput)
                .AddComponent("TestModule2", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindTestCleanupMethods(parser.State).ToList();

            Assert.AreEqual(1, initMethods.Count);
            Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedModuleName.ComponentName);
            Assert.AreEqual("TestCleanup", initMethods.ElementAt(0).MemberName);
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedTestInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput.Replace("'@TestInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindTestInitializeMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedTestCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput.Replace("'@TestCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindTestCleanupMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedTestInitInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetNormalModuleInput.Replace("'@TestInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindTestInitializeMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedTestCleanupInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetNormalModuleInput.Replace("'@TestCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindTestCleanupMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_DiscoversAnnotatedModuleInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput)
                .AddComponent("TestModule2", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindModuleInitializeMethods(parser.State).ToList();

            Assert.AreEqual(1, initMethods.Count);
            Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedModuleName.ComponentName);
            Assert.AreEqual("ModuleInitialize", initMethods.ElementAt(0).MemberName);
        }

        [TestMethod]
        public void Discovery_DiscoversAnnotatedModuleCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput)
                .AddComponent("TestModule2", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindModuleCleanupMethods(parser.State).ToList();

            Assert.AreEqual(1, initMethods.Count);
            Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedModuleName.ComponentName);
            Assert.AreEqual("ModuleCleanup", initMethods.ElementAt(0).MemberName);
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedModuleInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput.Replace("'@ModuleInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindModuleInitializeMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedModuleCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetTestModuleInput.Replace("'@ModuleCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindModuleCleanupMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedModuleInitInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetNormalModuleInput.Replace("'@ModuleInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindModuleInitializeMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
        }

        [TestMethod]
        public void Discovery_IgnoresNonAnnotatedModuleCleanupInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("TestModule1", vbext_ComponentType.vbext_ct_StdModule, GetNormalModuleInput.Replace("'@ModuleCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
            var qualifiedModuleName = new QualifiedModuleName(component);

            var initMethods = qualifiedModuleName.FindModuleCleanupMethods(parser.State);
            Assert.IsFalse(initMethods.Any());
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

        private string GetNormalModuleInput
        {
            get { return string.Format(RawInput, string.Empty); }
        }
    }
}