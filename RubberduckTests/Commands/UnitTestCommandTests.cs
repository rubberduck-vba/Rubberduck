using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Input;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.UnitTesting.ComCommands;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.UnitTesting.Settings;
using Rubberduck.VBEditor.Events;
using RubberduckTests.Settings;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class UnitTestCommandTests
    {
        private const string TestProjectName = "TestProject";
        private static readonly string TestModuleBaseName = TestExplorer.UnitTest_NewModule_BaseName;
        private static readonly string TestMethodBaseName = TestExplorer.UnitTest_NewMethod_BaseName;

        private const string TestModuleHeader = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
";

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddsTest(Type command)
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);

                addTestMethodCommand.Execute(null);

                var added = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test =>
                    test.Annotations.Any(pta => pta.Annotation is TestMethodAnnotation));

                Assert.NotNull(added);
            }
        }


        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand), 1, 2, 3)]
        [TestCase(typeof(AddTestMethodCommand), 1, 3, 2)]
        [TestCase(typeof(AddTestMethodCommand), 2, 3, 1)]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand), 1, 2, 3)]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand), 1, 3, 2)]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand), 2, 3, 1)]
        public void AddsTestPicksCorrectNumber(Type command, int first, int second, int added)
        {
            var input = $@"
{TestModuleHeader}
'@TestMethod
Public Sub {TestMethodBaseName}{first}()
End Sub
'@TestMethod
Public Sub {TestMethodBaseName}{second}()
End Sub
";

            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, input);
            using (state)
            {
                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);

                addTestMethodCommand.Execute(null);

                var expectedMethod = $"{TestMethodBaseName}{added}";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedMethod));

                Assert.NotNull(generated);
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddsTestPicksNextNumberAccountsForNonTests(Type command)
        {
            var input = $@"
{TestModuleHeader}
Public Function {TestMethodBaseName}1() As Long
End Function
'@TestMethod
Public Sub {TestMethodBaseName}2()
End Sub
";

            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, input);
            using (state)
            {
                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);

                addTestMethodCommand.Execute(null);

                var expectedMethod = $"{TestMethodBaseName}3";
                var added = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedMethod));

                Assert.NotNull(added);
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddsTest_NullActiveCodePane(Type command)
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);
                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);

                addTestMethodCommand.Execute(null);

                var unexpectedMethod = $"{TestMethodBaseName}3";
                var added = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(unexpectedMethod));

                Assert.IsNull(added);
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddTest_CanExecute_NonReadyState(Type command)
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, string.Empty);
            using (state)
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences, CancellationToken.None);

                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);

                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddTest_CanExecute(Type command)
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);
                Assert.IsTrue(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddTest_CanExecute_NoTestModule(Type command)
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, string.Empty);
            using (state)
            {
                var addTestMethodCommand = ArrangeAddTestMethodCommand(command, vbe, state);
                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModule()
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, "Module1", string.Empty);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state);

                addTestModuleCommand.Execute(null);

                var expectedModule = $"{TestModuleBaseName}1";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedModule));

                Assert.NotNull(generated);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleWithStubs()
        {
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, "Module1", string.Empty);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state);
                var addWithStubsCommand = ArrangeAddTestModuleWithStubsCommand(vbe, addTestModuleCommand);

                addWithStubsCommand.Execute(null);

                var expectedModule = $"{TestModuleBaseName}1";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedModule));

                Assert.NotNull(generated);
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(1, 2, 3)]
        [TestCase(1, 3, 2)]
        [TestCase(2, 3, 1)]
        public void AddsTestModulePicksCorrectNumber(int first, int second, int added)
        {
            var existing = new List<(ComponentType type, string name, string code)>
            {
                (ComponentType.StandardModule, $"{TestModuleBaseName}{first}", string.Empty),
                (ComponentType.StandardModule, $"{TestModuleBaseName}{second}", string.Empty)
            };

            var (vbe, state) = ArrangeAndParseTestCode(TestProjectName, existing);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state, ArrangeCodeGenerator(vbe.Object, state));

                addTestModuleCommand.Execute(null);

                var expectedModule = $"{TestModuleBaseName}{added}";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedModule));

                Assert.NotNull(generated);
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(1, 2, 3)]
        [TestCase(1, 3, 2)]
        [TestCase(2, 3, 1)]
        public void AddsTestModulePicksCorrectNumberWithStubs(int first, int second, int added)
        {
            var existing = new List<(ComponentType type, string name, string code)>
            {
                (ComponentType.StandardModule, $"{TestModuleBaseName}{first}", string.Empty),
                (ComponentType.StandardModule, $"{TestModuleBaseName}{second}", string.Empty)
            };

            var (vbe, state) = ArrangeAndParseTestCode(TestProjectName, existing);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state);
                var addWithStubsCommand = ArrangeAddTestModuleWithStubsCommand(vbe, addTestModuleCommand);
                List<Declaration> declarations = null;
                var model = new CodeExplorerComponentViewModel(null, null, ref declarations, vbe.Object);
                addWithStubsCommand.Execute(model);

                var expectedModule = $"{TestModuleBaseName}{added}";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedModule));

                Assert.NotNull(generated);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleWithStubsAddsStubsPublicProcedures()
        {
            const string code =
@"Public Sub PublicSub()
End Sub

Public Function PublicFunction()
End Function

Public Property Get PublicProperty()
End Property

Public Property Let PublicProperty(v As Variant)
End Property

Public Property Set PublicProperty(s As Object)
End Property
";
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, "Module1", code);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state);
                var addWithStubsCommand = ArrangeAddTestModuleWithStubsCommand(vbe, addTestModuleCommand);

                var project = state.DeclarationFinder.FindProject(TestProjectName);
                var target = state.DeclarationFinder.FindStdModule("Module1", project);

                var _ = Enumerable.Empty<Declaration>().ToList();
                var model = new CodeExplorerComponentViewModel(null, target, ref _, vbe.Object);

                addWithStubsCommand.Execute(model);

                var testModule = state.DeclarationFinder.FindStdModule($"{TestModuleBaseName}1", project);

                Assert.IsTrue(testModule.Annotations.Any(a => a.Annotation is TestModuleAnnotation));

                var stubIdentifierNames = new List<string>
                {
                    $"PublicSub{TestMethodBaseName}",
                    $"PublicFunction{TestMethodBaseName}",
                    $"GetPublicProperty{TestMethodBaseName}",
                    $"LetPublicProperty{TestMethodBaseName}",
                    $"SetPublicProperty{TestMethodBaseName}"
                };

                var stubs = state.DeclarationFinder.AllUserDeclarations.Where(d => d.IdentifierName.EndsWith(TestMethodBaseName)).ToList();

                Assert.AreEqual(stubIdentifierNames.Count, stubs.Count);
                Assert.IsTrue(stubs.All(d => stubIdentifierNames.Contains(d.IdentifierName)));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleWithStubsNoStubsAddedPrivateProcedures()
        {
            const string code =
@"Private Sub PrivateSub()
End Sub

Private Function PrivateFunction()
End Function

Private Property Get PrivateProperty()
End Property

Private Property Let PrivateProperty(v As Variant)
End Property

Private Property Set PrivateProperty(s As Object)
End Property
";
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, "Module1", code);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state, ArrangeCodeGenerator(vbe.Object, state));
                var addWithStubsCommand = ArrangeAddTestModuleWithStubsCommand(vbe, addTestModuleCommand);

                var project = state.DeclarationFinder.FindProject(TestProjectName);
                var target = state.DeclarationFinder.FindStdModule("Module1", project);

                var _ = Enumerable.Empty<Declaration>().ToList();
                var model = new CodeExplorerComponentViewModel(null, target, ref _, vbe.Object);

                addWithStubsCommand.Execute(model);

                var testModule = state.DeclarationFinder.FindStdModule($"{TestModuleBaseName}1", project);

                Assert.IsTrue(testModule.Annotations.Any(a => a.Annotation is TestModuleAnnotation));
                var stubs = state.DeclarationFinder.AllUserDeclarations.Where(d => d.IdentifierName.EndsWith(TestMethodBaseName)).ToList();

                Assert.AreEqual(0, stubs.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleWithStubsNoStubsAddedNonProcedures()
        {
            const string code =
@"Public Type UserDefinedType
    UserDefinedTypeMember As String
End Type

Public Declare PtrSafe Sub LibraryProcedure Lib ""lib.dll""()

Public Declare PtrSafe Function LibraryFunction Lib ""lib.dll""()

Public Variable As String

Public Const Constant As String = vbNullString

Public Enum Enumeration
    EnumerationMember
End Enum
";
            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, "Module1", code);
            using (state)
            {
                var addTestModuleCommand = ArrangeAddTestModuleCommand(vbe, state, ArrangeCodeGenerator(vbe.Object, state));
                var addWithStubsCommand = ArrangeAddTestModuleWithStubsCommand(vbe, addTestModuleCommand);

                var project = state.DeclarationFinder.FindProject(TestProjectName);
                var target = state.DeclarationFinder.FindStdModule("Module1", project);

                var _ = Enumerable.Empty<Declaration>().ToList();
                var model = new CodeExplorerComponentViewModel(null, target, ref _, vbe.Object);

                addWithStubsCommand.Execute(model);

                var testModule = state.DeclarationFinder.FindStdModule($"{TestModuleBaseName}1", project);

                Assert.IsTrue(testModule.Annotations.Any(a => a.Annotation is TestModuleAnnotation));
                var stubs = state.DeclarationFinder.AllUserDeclarations.Where(d => d.IdentifierName.EndsWith(TestMethodBaseName)).ToList();

                Assert.AreEqual(0, stubs.Count);
            }
        }

        // TODO: Remove the temporal copuling with other Arrange*
        private (Mock<IVBE> Vbe, RubberduckParserState State) ArrangeAndParseTestCode(ComponentType type, string name, string code)
        {
            return ArrangeAndParseTestCode(TestProjectName, new List<(ComponentType type, string name, string code)> {(type, name, code)});
        }

        // TODO: Remove the temporal copuling with other Arrange*
        private (Mock<IVBE> Vbe, RubberduckParserState State) ArrangeAndParseTestCode(string projectName, IEnumerable<(ComponentType type, string name, string code)> components)
        {
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected, ProjectType.StandAlone);

            foreach (var (type, name, code) in components)
            {
                projectBuilder.AddComponent(name, type, code);
            }

            var project = projectBuilder.Build();
            var vbe = builder.AddProject(project).Build();
            vbe.Setup(m => m.ActiveVBProject).Returns(project.Object);
            vbe.Setup(m => m.SelectedVBComponent).Returns(projectBuilder.MockVBComponents.Object.Last());

            var parser = MockParser.Create(vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(vbe));
            var state = parser.State;

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            return (vbe, state);
        }

        // TODO: Remove the temporal copuling with other Arrange*
        private AddTestModuleCommand ArrangeAddTestModuleCommand(Mock<IVBE> vbe, RubberduckParserState state)
        {
            return ArrangeAddTestModuleCommand(vbe, state, ArrangeCodeGenerator(vbe.Object, state));
        }

        // TODO: Remove the temporal copuling with other Arrange*
        private AddTestModuleCommand ArrangeAddTestModuleCommand(Mock<IVBE> vbe, RubberduckParserState state, ITestCodeGenerator generator)
        {
            return ArrangeAddTestModuleCommand(vbe, state, generator, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        // TODO: Remove the temporal copuling with other Arrange*
        private AddTestModuleCommand ArrangeAddTestModuleCommand(Mock<IVBE> vbe, RubberduckParserState state, ITestCodeGenerator generator, Mock<IVbeEvents> vbeEvents)
        {
            return new AddTestModuleCommand(vbe.Object, state, ArrangeCodeGenerator(vbe.Object, state), vbeEvents.Object, state.ProjectsProvider);
        }

        // TODO: Remove the temporal coupling with other Arrange*
        private AddTestModuleWithStubsCommand ArrangeAddTestModuleWithStubsCommand(Mock<IVBE> vbe,
            AddTestModuleCommand addTestModuleCommand)
        {
            return ArrangeAddTestModuleWithStubsCommand(vbe, addTestModuleCommand, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        // TODO: Remove the temporal copuling with other Arrange*
        private AddTestModuleWithStubsCommand ArrangeAddTestModuleWithStubsCommand(Mock<IVBE> vbe, AddTestModuleCommand addTestModuleCommand, Mock<IVbeEvents> vbeEvents)
        {
            return new AddTestModuleWithStubsCommand(vbe.Object, addTestModuleCommand, vbeEvents.Object);
        }

        // TODO: Remove the temporal coupling with other Arrange*
        private ICommand ArrangeAddTestMethodCommand(Type command, Mock<IVBE> vbe, RubberduckParserState state)
        {
            return ArrangeAddTestMethodCommand(command, vbe, state, ArrangeCodeGenerator(vbe.Object, state));
        }

        // TODO: Remove the temporal coupling with other Arrange*
        private ICommand ArrangeAddTestMethodCommand(Type command, Mock<IVBE> vbe, RubberduckParserState state,
            ITestCodeGenerator testCodeGenerator)
        {
            return ArrangeAddTestMethodCommand(command, vbe, state, testCodeGenerator, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        // TODO: Remove the temporal coupling with other Arrange*
        private ICommand ArrangeAddTestMethodCommand(Type command, Mock<IVBE> vbe, RubberduckParserState state,
            ITestCodeGenerator testCodeGenerator, Mock<IVbeEvents> vbeEvents)
        {
            return (ICommand) Activator.CreateInstance(command, vbe.Object, state, testCodeGenerator, vbeEvents.Object);
        }

        private ITestCodeGenerator ArrangeCodeGenerator(IVBE vbe, RubberduckParserState state)
        {
            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var settings = new Mock<IConfigurationService<UnitTestSettings>>();
            var arguments = new Mock<ICompilationArgumentsProvider>();

            settings.Setup(s => s.Read()).Returns(new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, true, true, false));
            arguments.Setup(m => m.UserDefinedCompilationArguments(It.IsAny<string>()))
                .Returns(new Dictionary<string, short>());

            return new TestCodeGenerator(vbe, state, new Mock<IMessageBox>().Object, new Mock<IVBEInteraction>().Object, settings.Object, indenter, arguments.Object);           
        }
    }
}