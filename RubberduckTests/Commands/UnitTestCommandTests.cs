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
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.Utility;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.ComManagement;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class UnitTestCommandTests
    {
        private struct AddTestModuleCommandTestState
        {
            public Mock<IVBE> Vbe;
            public AddTestModuleCommand AddTestModuleCommand;
            public AddTestModuleWithStubsCommand AddTestModuleWithStubsCommand;
        }

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
            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                addTestMethodCommand.Execute(null);

                var added = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test =>
                    test.Annotations.Any(pta => pta.Annotation is TestMethodAnnotation));

                Assert.NotNull(added);
            }
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/4169
        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddsTestTBD(Type command)
        {
            var inputCode =
@"
Option Explicit
Option Private Module

'@TestModule
'@Folder(""Tests"")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject(""Rubberduck.AssertClass"")
    Set Fakes = CreateObject(""Rubberduck.FakesProvider"")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
";
        var(vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, inputCode);
            using (state)
            {
                addTestMethodCommand.Execute(null);

                addTestMethodCommand.Execute(null);

                var expectedMethod = $"{TestMethodBaseName}{1}";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedMethod));

                var lastExistingProcedure = "TestCleanup";
                var lastProc = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(lastExistingProcedure));

                Assert.IsTrue(lastProc.Context.Stop.Line < generated.Context.Start.Line);
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

            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, input);
            using (state)
            {
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

            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, input);
            using (state)
            {
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
            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);

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
            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences, CancellationToken.None);

                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddTest_CanExecute(Type command)
        {
            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, TestModuleHeader);
            using (state)
            {
                Assert.IsTrue(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddTest_CanExecute_NoTestModule(Type command)
        {
            var (vbe, state, addTestMethodCommand) = ParseAndArrangeAddTestMethodCommandTests(command, ComponentType.StandardModule, TestModuleBaseName, string.Empty);
            using (state)
            {
                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModule()
        {
            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(ComponentType.StandardModule, "Module1", string.Empty);
            using (state)
            {

                testState.AddTestModuleCommand.Execute(null);

                var expectedModule = $"{TestModuleBaseName}1";
                var generated = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedModule));

                Assert.NotNull(generated);
            }
        }

        [Category("Commands")]
        [Test]
        public void AddsTestModuleWithStubs()
        {
            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(ComponentType.StandardModule, "Module1", string.Empty);
            using (state)
            {
                testState.AddTestModuleWithStubsCommand.Execute(null);

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

            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(TestProjectName, existing);
            using (state)
            {
                testState.AddTestModuleCommand.Execute(null);

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

            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(TestProjectName, existing);
            using (state)
            {
                List<Declaration> declarations = null;
                var model = new CodeExplorerComponentViewModel(null, null, ref declarations,testState.Vbe.Object);
                testState.AddTestModuleWithStubsCommand.Execute(model);

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
            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(ComponentType.StandardModule, "Module1", code);
            using (state)
            {
                var project = state.DeclarationFinder.FindProject(TestProjectName);
                var target = state.DeclarationFinder.FindStdModule("Module1", project);

                var _ = Enumerable.Empty<Declaration>().ToList();
                var model = new CodeExplorerComponentViewModel(null, target, ref _, testState.Vbe.Object);

                testState.AddTestModuleWithStubsCommand.Execute(model);

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
            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(ComponentType.StandardModule, "Module1", code);
            using (state)
            {
                var project = state.DeclarationFinder.FindProject(TestProjectName);
                var target = state.DeclarationFinder.FindStdModule("Module1", project);

                var _ = Enumerable.Empty<Declaration>().ToList();
                var model = new CodeExplorerComponentViewModel(null, target, ref _, testState.Vbe.Object);

                testState.AddTestModuleWithStubsCommand.Execute(model);

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
            var (testState, state) = ParseAndArrangeAddTestModuleCommandTests(ComponentType.StandardModule, "Module1", code);
            using (state)
            {
                var project = state.DeclarationFinder.FindProject(TestProjectName);
                var target = state.DeclarationFinder.FindStdModule("Module1", project);

                var _ = Enumerable.Empty<Declaration>().ToList();
                var model = new CodeExplorerComponentViewModel(null, target, ref _, testState.Vbe.Object);

                testState.AddTestModuleWithStubsCommand.Execute(model);

                var testModule = state.DeclarationFinder.FindStdModule($"{TestModuleBaseName}1", project);

                Assert.IsTrue(testModule.Annotations.Any(a => a.Annotation is TestModuleAnnotation));
                var stubs = state.DeclarationFinder.AllUserDeclarations.Where(d => d.IdentifierName.EndsWith(TestMethodBaseName)).ToList();

                Assert.AreEqual(0, stubs.Count);
            }
        }

        private (Mock<IVBE> Vbe, RubberduckParserState State, ICommand command) ParseAndArrangeAddTestMethodCommandTests(Type command, ComponentType type, string name, string code)
        {
            var components = new List<(ComponentType type, string name, string code)> { (type, name, code) };
            return ParseAndArrangeAddTestMethodCommandTests(command, TestProjectName, components);
        }

        private (Mock<IVBE> Vbe, RubberduckParserState State, ICommand command) ParseAndArrangeAddTestMethodCommandTests(Type command, string projectName, IEnumerable<(ComponentType type, string name, string code)> components)
        {
            var vbe = BuildMockVBE(projectName, components);

            (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) = CreateAndParseWithRewritingManager(vbe);
            var state = parser.State;

            var testCodeGenerator = ArrangeCodeGenerator(vbe.Object, state);
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var cmd = (ICommand)Activator.CreateInstance(command, vbe.Object, state, rewritingManager, testCodeGenerator, vbeEvents.Object);
            return (vbe, state, cmd);

        }

        private (AddTestModuleCommandTestState testState, RubberduckParserState state)  ParseAndArrangeAddTestModuleCommandTests(ComponentType type, string name, string code)
        {
            var components = new List<(ComponentType type, string name, string code)> { (type, name, code) };
            return ParseAndArrangeAddTestModuleCommandTests(TestProjectName, components);
        }

        private (AddTestModuleCommandTestState testState, RubberduckParserState state)  ParseAndArrangeAddTestModuleCommandTests(string projectName, IEnumerable<(ComponentType type, string name, string code)> components)
        {
            var vbe = BuildMockVBE(projectName, components);

            (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) = CreateAndParseWithRewritingManager(vbe);

            var state = parser.State;

            var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, ArrangeCodeGenerator(vbe.Object, state), MockVbeEvents.CreateMockVbeEvents(vbe).Object, state.ProjectsProvider);
            var addWithStubsCommand = new AddTestModuleWithStubsCommand(vbe.Object, addTestModuleCommand, MockVbeEvents.CreateMockVbeEvents(vbe).Object);

            var testStatus = new AddTestModuleCommandTestState();
            testStatus.Vbe = vbe;
            testStatus.AddTestModuleCommand = addTestModuleCommand;
            testStatus.AddTestModuleWithStubsCommand = addWithStubsCommand;

            return (testStatus, state);

        }

        private (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) CreateAndParseWithRewritingManager(Mock<IVBE> vbe)
        {
            (SynchronousParseCoordinator parser, IRewritingManager rewritingManager) = MockParser.CreateWithRewriteManager(vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(vbe));
           
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            return (parser, rewritingManager);
        }

        private static Mock<IVBE> BuildMockVBE(string projectName, IEnumerable<(ComponentType type, string name, string code)> components)
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
            return vbe;
        }

        private static ITestCodeGenerator ArrangeCodeGenerator(IVBE vbe, RubberduckParserState state)
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