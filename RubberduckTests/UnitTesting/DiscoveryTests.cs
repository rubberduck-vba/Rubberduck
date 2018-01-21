using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    public class DiscoveryTests
    {
        [Test]
        [Category("Unit Testing")]
        public void Discovery_DiscoversAnnotatedTestMethods()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                Assert.AreEqual(1, UnitTestUtils.GetAllTests(vbe, state).Count());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedTestMethods()
        {
            var testMethods = @"Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                Assert.IsFalse(UnitTestUtils.GetAllTests(vbe, state).Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresAnnotatedTestMethodsNotInTestModule()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                Assert.IsFalse(UnitTestUtils.GetAllTests(vbe, state).Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_DiscoversAnnotatedTestMethodsInGivenTestModule()
        {
            var testMethods = @"'@TestMethod
Public Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput + testMethods)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var tests = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object.GetTests(vbe, state).ToList();

                Assert.AreEqual(1, tests.Count);
                Assert.AreEqual("TestModule1", tests.ElementAt(0).Declaration.ComponentName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_DiscoversAnnotatedTestInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindTestInitializeMethods(state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("TestInitialize", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_DiscoversAnnotatedTestCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindTestCleanupMethods(state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("TestCleanup", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedTestInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput.Replace("'@TestInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindTestInitializeMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedTestCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput.Replace("'@TestCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindTestCleanupMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedTestInitInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput.Replace("'@TestInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindTestInitializeMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedTestCleanupInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput.Replace("'@TestCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindTestCleanupMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_DiscoversAnnotatedModuleInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindModuleInitializeMethods(state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("ModuleInitialize", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_DiscoversAnnotatedModuleCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindModuleCleanupMethods(state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("ModuleCleanup", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedModuleInitInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput.Replace("'@ModuleInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindModuleInitializeMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedModuleCleanupInGivenTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput.Replace("'@ModuleCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindModuleCleanupMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedModuleInitInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput.Replace("'@ModuleInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindModuleInitializeMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void Discovery_IgnoresNonAnnotatedModuleCleanupInGivenNonTestModule()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput.Replace("'@ModuleCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = qualifiedModuleName.FindModuleCleanupMethods(state);
                Assert.IsFalse(initMethods.Any());
            }
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