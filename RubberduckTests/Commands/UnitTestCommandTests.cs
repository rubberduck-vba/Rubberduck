using System;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UnitTesting;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class UnitTestCommandTests
    {
        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(
                    string.Format(input,
                        AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder, "TestMethod1")) +
                    Environment.NewLine, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestPicksNextNumber()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Sub TestMethod1()
End Sub
'@TestMethod
Public Sub TestMethod2()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                var expected = string.Format(input,
                                   AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder,
                                       "TestMethod3")) +
                               Environment.NewLine;
                var actual = module.Content();
                Assert.AreEqual(expected, actual);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestPicksNextNumberAccountsForNonTests()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
Public Function TestMethod1() As Long
End Function
'@TestMethod
Public Sub TestMethod2()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                var expected = string.Format(input,
                                   AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder,
                                       "TestMethod3")) +
                               Environment.NewLine;
                var actual = module.Content();
                Assert.AreEqual(expected, actual);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestPicksNextNumberGapExists()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Sub TestMethod1()
End Sub
'@TestMethod
Public Sub TestMethod3()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                var expected = string.Format(input,
                                   AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder,
                                       "TestMethod2")) +
                               Environment.NewLine;
                var actual = module.Content();
                Assert.AreEqual(expected, actual);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestPicksNextNumberGapAtStart()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Sub TestMethod2()
End Sub
'@TestMethod
Public Sub TestMethod3()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                var expected = string.Format(input,
                                   AddTestMethodCommand.TestMethodTemplate.Replace(AddTestMethodCommand.NamePlaceholder,
                                       "TestMethod1")) +
                               Environment.NewLine;
                var actual = module.Content();
                Assert.AreEqual(expected, actual);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsExpectedErrorTestPicksNextNumber()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Sub TestMethod1()
End Sub
'@TestMethod
Public Sub TestMethod2()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(
                    string.Format(input,
                        AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
                            "TestMethod3")) + Environment.NewLine, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsExpectedErrorTestPicksNextNumberAccountsForNonTests()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Function TestMethod1() As Long
End Function
'@TestMethod
Public Sub TestMethod2()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(
                    string.Format(input,
                        AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
                            "TestMethod3")) + Environment.NewLine, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsExpectedErrorTestPicksNextNumberGapExists()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Sub TestMethod1()
End Sub
'@TestMethod
Public Sub TestMethod3()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(
                    string.Format(input,
                        AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
                            "TestMethod2")) + Environment.NewLine, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsExpectedErrorTestPicksNextNumberGapAtStart()
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
'@TestMethod
Public Sub TestMethod2()
End Sub
'@TestMethod
Public Sub TestMethod3()
End Sub
{0}";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(
                    string.Format(input,
                        AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
                            "TestMethod1")) + Environment.NewLine, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(input, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddTest_CanExecute_NonReadyState()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences, CancellationToken.None);

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddTest_CanExecute_NoTestModule()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
                Assert.IsTrue(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

                addTestMethodCommand.Execute(null);
                var module = component.CodeModule;

                Assert.AreEqual(
                    string.Format(input,
                        AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
                            "TestMethod1")) + Environment.NewLine, module.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddExpectedErrorTest_CanExecute_NonReadyState()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences, CancellationToken.None);

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddExpectedErrorTest_CanExecute_NoTestModule()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
                Assert.IsTrue(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
                addTestMethodCommand.Execute(null);

                Assert.AreEqual(input, component.CodeModule.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModule()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var interaction = new Mock<IVBEInteraction>();
                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
                var config = GetUnitTestConfig();
                settings.Setup(x => x.LoadConfiguration()).Returns(config);


                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
                addTestModuleCommand.Execute(null);

                // mock suite auto-assigns "TestModule1" to the first component when we create the mock
                var project = state.DeclarationFinder.FindProject("TestProject1");
                var module = state.DeclarationFinder.FindStdModule("TestModule2", project);
                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleNextAvailableNumberGapInSequence()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, string.Empty)
                .AddComponent("TestModule3", ComponentType.StandardModule, string.Empty)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var interaction = new Mock<IVBEInteraction>();
                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
                var config = GetUnitTestConfig();
                settings.Setup(x => x.LoadConfiguration()).Returns(config);

                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
                addTestModuleCommand.Execute(null);

                var declaration = state.DeclarationFinder.FindProject("TestProject1");
                var module = state.DeclarationFinder.FindStdModule("TestModule2", declaration);
                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleNextAvailableNumberGapAtStart()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule2", ComponentType.StandardModule, string.Empty)
                .AddComponent("TestModule3", ComponentType.StandardModule, string.Empty)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var interaction = new Mock<IVBEInteraction>();
                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
                var config = GetUnitTestConfig();
                settings.Setup(x => x.LoadConfiguration()).Returns(config);

                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
                addTestModuleCommand.Execute(null);

                var declaration = state.DeclarationFinder.FindProject("TestProject1");
                var module = state.DeclarationFinder.FindStdModule("TestModule1", declaration);
                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleNextAvailableNumberNoGaps()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, string.Empty)
                .AddComponent("TestModule2", ComponentType.StandardModule, string.Empty)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var interaction = new Mock<IVBEInteraction>();
                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
                var config = GetUnitTestConfig();
                settings.Setup(x => x.LoadConfiguration()).Returns(config);

                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
                addTestModuleCommand.Execute(null);

                var declaration = state.DeclarationFinder.FindProject("TestProject1");
                var module = state.DeclarationFinder.FindStdModule("TestModule3", declaration);
                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleWithStubs()
        {
            const string code =
                @"Public Type UserDefinedType
    UserDefinedTypeMember As String
End Type

Public Declare PtrSafe Sub LibraryProcedure Lib ""lib.dll"" ()

Public Declare PtrSafe Function LibraryFunction Lib ""lib.dll"" ()

Public Variable As String

Public Const Constant As String = """"

Public Enum Enumeration
    EnumerationMember
End Enum

Public Sub PublicProcedure(Parameter As String)
    Dim LocalVariable as String
    Const LocalConstant as String = """"
LineLabel:
End Sub

Public Function PublicFunction()
End Function

Public Property Get PublicProperty()
End Property

Public Property Let PublicProperty(v As Variant)
End Property

Public Property Set PublicProperty(s As String)
End Property

Private Sub PrivateProcedure(Parameter As String)
End Sub

Private Function PrivateFunction()
End Function

Private Property Get PrivateProperty()
End Property

Private Property Let PrivateProperty(v As Variant)
End Property

Private Property Set PrivateProperty(s As String)
End Property";
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
                var interaction = new Mock<IVBEInteraction>();
                var config = GetUnitTestConfig();
                settings.Setup(x => x.LoadConfiguration()).Returns(config);

                var project = state.DeclarationFinder.FindProject("TestProject1");
                var module = state.DeclarationFinder.FindStdModule("TestModule1", project);

                var messageBox = new Mock<IMessageBox>();
                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
                addTestModuleCommand.Execute(module);

                var testModule = state.DeclarationFinder.FindStdModule("TestModule2", project);

                var stubIdentifierNames = new[]
                {
                    "PublicProcedure_Test", "PublicFunction_Test", "GetPublicProperty_Test", "LetPublicProperty_Test", "SetPublicProperty_Test"
                };

                Assert.IsTrue(testModule.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));

                var stubs = state.DeclarationFinder.AllUserDeclarations.Where(d => d.IdentifierName.EndsWith("_Test")).ToList();

                Assert.AreEqual(stubIdentifierNames.Length, stubs.Count);
                Assert.IsTrue(stubs.All(d => stubIdentifierNames.Contains(d.IdentifierName)));
            }
        }

        private Configuration GetUnitTestConfig()
        {
            var unitTestSettings = new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, false, false, false);

            var userSettings = new UserSettings(null, null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }
    }
}