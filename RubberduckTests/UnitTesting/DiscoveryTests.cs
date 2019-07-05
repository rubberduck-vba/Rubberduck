using System.Linq;
using NUnit.Framework;
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
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_DiscoversAnnotatedTestMethods(string accessibility)
        {
            var testMethods = $@"'@TestMethod
{accessibility} Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility) + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                Assert.AreEqual(1, TestDiscovery.GetAllTests(state).Count());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedTestMethods(string accessibility)
        {
            var testMethods = $@"{accessibility} Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility) + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                Assert.IsFalse(TestDiscovery.GetAllTests(state).Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresAnnotatedTestMethodsNotInTestModule(string accessibility)
        {
            var testMethods = $@"'@TestMethod
{accessibility} Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput(accessibility) + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                Assert.IsFalse(TestDiscovery.GetAllTests(state).Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_DiscoversAnnotatedTestMethodsInGivenTestModule(string accessibility)
        {
            var testMethods = $@"'@TestMethod
{accessibility} Sub TestMethod1()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility) + testMethods)
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput(accessibility) + testMethods);

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var tests = TestDiscovery.GetTests(vbe, component, state).ToList();

                Assert.AreEqual(1, tests.Count);
                Assert.AreEqual("TestModule1", tests.ElementAt(0).Declaration.ComponentName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_DiscoversAnnotatedTestInitInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility))
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput(accessibility));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);
                
                var initMethods = TestDiscovery.FindTestInitializeMethods(qualifiedModuleName, state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("TestInitialize", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_DiscoversAnnotatedTestCleanupInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility))
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput(accessibility));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindTestCleanupMethods(qualifiedModuleName, state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("TestCleanup", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedTestInitInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility).Replace("'@TestInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindTestInitializeMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedTestCleanupInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility).Replace("'@TestCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindTestCleanupMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedTestInitInGivenNonTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput(accessibility).Replace("'@TestInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindTestInitializeMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedTestCleanupInGivenNonTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput(accessibility).Replace("'@TestCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindTestCleanupMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_DiscoversAnnotatedModuleInitInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility))
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput(accessibility));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindModuleInitializeMethods(qualifiedModuleName, state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("ModuleInitialize", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_DiscoversAnnotatedModuleCleanupInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility))
                .AddComponent("TestModule2", ComponentType.StandardModule, GetTestModuleInput(accessibility));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindModuleCleanupMethods(qualifiedModuleName, state).ToList();

                Assert.AreEqual(1, initMethods.Count);
                Assert.AreEqual("TestModule1", initMethods.ElementAt(0).QualifiedName.QualifiedModuleName.ComponentName);
                Assert.AreEqual("ModuleCleanup", initMethods.ElementAt(0).QualifiedName.MemberName);
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedModuleInitInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility).Replace("'@ModuleInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindModuleInitializeMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedModuleCleanupInGivenTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetTestModuleInput(accessibility).Replace("'@ModuleCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindModuleCleanupMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedModuleInitInGivenNonTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput(accessibility).Replace("'@ModuleInitialize", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindModuleInitializeMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        [Test]
        [Category("Unit Testing")]
        [TestCase("Public")]
        [TestCase("Private")]
        [TestCase("Friend")]
        public void Discovery_IgnoresNonAnnotatedModuleCleanupInGivenNonTestModule(string accessibility)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, GetNormalModuleInput(accessibility).Replace("'@ModuleCleanup", string.Empty));

            var vbe = builder.AddProject(project.Build()).Build().Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var component = project.MockComponents.Single(f => f.Object.Name == "TestModule1").Object;
                var qualifiedModuleName = new QualifiedModuleName(component);

                var initMethods = TestDiscovery.FindModuleCleanupMethods(qualifiedModuleName, state);
                Assert.IsFalse(initMethods.Any());
            }
        }

        private const string RawInput = @"Option Explicit
Option Private Module

{0}

Private Assert As New Rubberduck.AssertClass

'@ModuleInitialize
{1} Sub ModuleInitialize()
    'this method runs once per module.
End Sub

'@ModuleCleanup
{1} Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
{1} Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
{1} Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
";

        private static string GetTestModuleInput(string accessibility)
        {
            return string.Format(RawInput, "'@TestModule", accessibility);
        }

        private static string GetNormalModuleInput(string accessibility)
        {
            return string.Format(RawInput, string.Empty, accessibility);
        }
    }
}