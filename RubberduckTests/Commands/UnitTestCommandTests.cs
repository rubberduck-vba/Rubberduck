using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Input;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.UnitTesting.Settings;
using RubberduckTests.Settings;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class UnitTestCommandTests
    {
        private static readonly string TestModuleBaseName = TestExplorer.UnitTest_NewModule_BaseName;
        private static readonly string TestMethodBaseName = TestExplorer.UnitTest_NewMethod_BaseName;

        [Category("Commands")]
        [Test]
        [TestCase(typeof(AddTestMethodCommand))]
        [TestCase(typeof(AddTestMethodExpectedErrorCommand))]
        public void AddsTest(Type command)
        {
            var input = @"
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
";

            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, input);
            using (state)
            {
                var addTestMethodCommand = (ICommand)Activator.CreateInstance(command, vbe, state, ArrangeCodeGenerator(vbe, state));

                addTestMethodCommand.Execute(null);

                var added = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test =>
                    test.Annotations.Any(annotation => annotation is TestMethodAnnotation));

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
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
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
                var addTestMethodCommand = (ICommand)Activator.CreateInstance(command, vbe, state, ArrangeCodeGenerator(vbe, state));

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
Option Explicit
Option Private Module

'@TestModule

Private Assert As Object
Public Function {TestMethodBaseName}1() As Long
End Function
'@TestMethod
Public Sub {TestMethodBaseName}2()
End Sub
";

            var (vbe, state) = ArrangeAndParseTestCode(ComponentType.StandardModule, TestModuleBaseName, input);
            using (state)
            {
                var addTestMethodCommand = (ICommand)Activator.CreateInstance(command, vbe, state, ArrangeCodeGenerator(vbe, state));

                addTestMethodCommand.Execute(null);

                var expectedMethod = $"{TestMethodBaseName}3";
                var added = state.DeclarationFinder.AllUserDeclarations.SingleOrDefault(test => test.IdentifierName.Equals(expectedMethod));

                Assert.NotNull(added);
            }
        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsTest_NullActiveCodePane()
        //        {
        //            var input = @"
        //Option Explicit
        //Option Private Module

        //'@TestModule
        //Private Assert As Object
        //";

        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
        //            vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);

        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);

        //                addTestMethodCommand.Execute(null);
        //                var module = component.CodeModule;

        //                Assert.AreEqual(input, module.Content());
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddTest_CanExecute_NonReadyState()
        //        {
        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences, CancellationToken.None);

        //                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
        //                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddTest_CanExecute_NoTestModule()
        //        {
        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
        //                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddTest_CanExecute()
        //        {
        //            var input = @"
        //Option Explicit
        //Option Private Module

        //'@TestModule

        //Private Assert As Object
        //";

        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodCommand(vbe.Object, state);
        //                Assert.IsTrue(addTestMethodCommand.CanExecute(null));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsExpectedErrorTest()
        //        {
        //            var input = @"
        //Option Explicit
        //Option Private Module

        //'@TestModule

        //Private Assert As Object
        //{0}";

        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(input, string.Empty), out component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);

        //                addTestMethodCommand.Execute(null);
        //                var module = component.CodeModule;

        //                Assert.AreEqual(
        //                    string.Format(input,
        //                        AddTestMethodExpectedErrorCommand.TestMethodExpectedErrorTemplate.Replace(AddTestMethodExpectedErrorCommand.NamePlaceholder,
        //                            "TestMethod1")) + Environment.NewLine, module.Content());
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddExpectedErrorTest_CanExecute_NonReadyState()
        //        {
        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);

        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {
        //                state.SetStatusAndFireStateChanged(this, ParserState.ResolvingReferences, CancellationToken.None);

        //                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
        //                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddExpectedErrorTest_CanExecute_NoTestModule()
        //        {
        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
        //                Assert.IsFalse(addTestMethodCommand.CanExecute(null));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddExpectedErrorTest_CanExecute()
        //        {
        //            var input = @"
        //Option Explicit
        //Option Private Module

        //'@TestModule

        //Private Assert As Object
        //";

        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
        //                Assert.IsTrue(addTestMethodCommand.CanExecute(null));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsExpectedErrorTest_NullActiveCodePane()
        //        {
        //            var input = @"
        //Option Explicit
        //Option Private Module

        //'@TestModule
        //Private Assert As Object
        //";

        //            IVBComponent component;
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
        //            vbe.Setup(s => s.ActiveCodePane).Returns((ICodePane)null);

        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {

        //                var addTestMethodCommand = new AddTestMethodExpectedErrorCommand(vbe.Object, state);
        //                addTestMethodCommand.Execute(null);

        //                Assert.AreEqual(input, component.CodeModule.Content());
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsTestModule()
        //        {
        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out var component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {
        //                var messageBox = new Mock<IMessageBox>();
        //                var interaction = new Mock<IVBEInteraction>();
        //                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
        //                var config = GetUnitTestConfig();
        //                settings.Setup(x => x.LoadConfiguration()).Returns(config);


        //                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
        //                addTestModuleCommand.Execute(null);

        //                // mock suite auto-assigns "TestModule1" to the first component when we create the mock
        //                var project = state.DeclarationFinder.FindProject("TestProject1");
        //                var module = state.DeclarationFinder.FindStdModule("TestModule2", project);
        //                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsTestModuleNextAvailableNumberGapInSequence()
        //        {
        //            var builder = new MockVbeBuilder();
        //            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
        //                .AddComponent("TestModule1", ComponentType.StandardModule, string.Empty)
        //                .AddComponent("TestModule3", ComponentType.StandardModule, string.Empty)
        //                .Build();
        //            var vbe = builder.AddProject(project).Build();

        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {
        //                var messageBox = new Mock<IMessageBox>();
        //                var interaction = new Mock<IVBEInteraction>();
        //                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
        //                var config = GetUnitTestConfig();
        //                settings.Setup(x => x.LoadConfiguration()).Returns(config);

        //                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
        //                addTestModuleCommand.Execute(null);

        //                var declaration = state.DeclarationFinder.FindProject("TestProject1");
        //                var module = state.DeclarationFinder.FindStdModule("TestModule2", declaration);
        //                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsTestModuleNextAvailableNumberGapAtStart()
        //        {
        //            var builder = new MockVbeBuilder();
        //            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
        //                .AddComponent("TestModule2", ComponentType.StandardModule, string.Empty)
        //                .AddComponent("TestModule3", ComponentType.StandardModule, string.Empty)
        //                .Build();
        //            var vbe = builder.AddProject(project).Build();

        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {
        //                var messageBox = new Mock<IMessageBox>();
        //                var interaction = new Mock<IVBEInteraction>();
        //                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
        //                var config = GetUnitTestConfig();
        //                settings.Setup(x => x.LoadConfiguration()).Returns(config);

        //                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
        //                addTestModuleCommand.Execute(null);

        //                var declaration = state.DeclarationFinder.FindProject("TestProject1");
        //                var module = state.DeclarationFinder.FindStdModule("TestModule1", declaration);
        //                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsTestModuleNextAvailableNumberNoGaps()
        //        {
        //            var builder = new MockVbeBuilder();
        //            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
        //                .AddComponent("TestModule1", ComponentType.StandardModule, string.Empty)
        //                .AddComponent("TestModule2", ComponentType.StandardModule, string.Empty)
        //                .Build();
        //            var vbe = builder.AddProject(project).Build();

        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {
        //                var messageBox = new Mock<IMessageBox>();
        //                var interaction = new Mock<IVBEInteraction>();
        //                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
        //                var config = GetUnitTestConfig();
        //                settings.Setup(x => x.LoadConfiguration()).Returns(config);

        //                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
        //                addTestModuleCommand.Execute(null);

        //                var declaration = state.DeclarationFinder.FindProject("TestProject1");
        //                var module = state.DeclarationFinder.FindStdModule("TestModule3", declaration);
        //                Assert.IsTrue(module.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));
        //            }
        //        }

        //        [Category("Commands")]
        //        [Test]
        //        public void AddsTestModuleWithStubs()
        //        {
        //            const string code =
        //                @"Public Type UserDefinedType
        //    UserDefinedTypeMember As String
        //End Type

        //Public Declare PtrSafe Sub LibraryProcedure Lib ""lib.dll"" ()

        //Public Declare PtrSafe Function LibraryFunction Lib ""lib.dll"" ()

        //Public Variable As String

        //Public Const Constant As String = """"

        //Public Enum Enumeration
        //    EnumerationMember
        //End Enum

        //Public Sub PublicProcedure(Parameter As String)
        //    Dim LocalVariable as String
        //    Const LocalConstant as String = """"
        //LineLabel:
        //End Sub

        //Public Function PublicFunction()
        //End Function

        //Public Property Get PublicProperty()
        //End Property

        //Public Property Let PublicProperty(v As Variant)
        //End Property

        //Public Property Set PublicProperty(s As String)
        //End Property

        //Private Sub PrivateProcedure(Parameter As String)
        //End Sub

        //Private Function PrivateFunction()
        //End Function

        //Private Property Get PrivateProperty()
        //End Property

        //Private Property Let PrivateProperty(v As Variant)
        //End Property

        //Private Property Set PrivateProperty(s As String)
        //End Property";

        //            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out var component);
        //            using (var state = MockParser.CreateAndParse(vbe.Object))
        //            {
        //                var settings = new Mock<ConfigurationLoader>(null, null, null, null, null, null, null, null);
        //                var interaction = new Mock<IVBEInteraction>();
        //                var config = GetUnitTestConfig();
        //                settings.Setup(x => x.LoadConfiguration()).Returns(config);

        //                var project = state.DeclarationFinder.FindProject("TestProject1");
        //                var module = state.DeclarationFinder.FindStdModule("TestModule1", project);

        //                var messageBox = new Mock<IMessageBox>();
        //                var addTestModuleCommand = new AddTestModuleCommand(vbe.Object, state, settings.Object, messageBox.Object, interaction.Object);
        //                addTestModuleCommand.Execute(module);

        //                var testModule = state.DeclarationFinder.FindStdModule("TestModule2", project);

        //                var stubIdentifierNames = new[]
        //                {
        //                    "PublicProcedure_Test", "PublicFunction_Test", "GetPublicProperty_Test", "LetPublicProperty_Test", "SetPublicProperty_Test"
        //                };

        //                Assert.IsTrue(testModule.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));

        //                var stubs = state.DeclarationFinder.AllUserDeclarations.Where(d => d.IdentifierName.EndsWith("_Test")).ToList();

        //                Assert.AreEqual(stubIdentifierNames.Length, stubs.Count);
        //                Assert.IsTrue(stubs.All(d => stubIdentifierNames.Contains(d.IdentifierName)));
        //            }
        //        }

        private (IVBE Vbe, RubberduckParserState State) ArrangeAndParseTestCode(ComponentType type, string name, string code)
        {
            return ArrangeAndParseTestCode("TestProject",
                new List<(ComponentType type, string name, string code)> {(type, name, code)});
        }

        private (IVBE Vbe, RubberduckParserState State) ArrangeAndParseTestCode(string projectName, List<(ComponentType type, string name, string code)> components)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected, ProjectType.StandAlone);

            foreach (var (type, name, code) in components)
            {
                project.AddComponent(name, type, code);
            }

            var vbe = builder.AddProject(project.Build()).Build();
            vbe.Setup(m => m.SelectedVBComponent).Returns(project.MockVBComponents.Object.Last());

            var parser = MockParser.Create(vbe.Object, null, MockVbeEvents.CreateMockVbeEvents(vbe));
            var state = parser.State;

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            return (vbe.Object, state);
        }

        private ITestCodeGenerator ArrangeCodeGenerator(IVBE vbe, RubberduckParserState state)
        {
            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var settings = new Mock<IConfigProvider<UnitTestSettings>>();

            settings.Setup(s => s.Create()).Returns(new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, true, true, false));

            return new TestCodeGenerator(vbe, state, new Mock<IMessageBox>().Object, new Mock<IVBEInteraction>().Object, settings.Object, indenter);           
        }

        private Configuration GetUnitTestConfig()
        {
            var unitTestSettings = new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, false, false, false);

            var userSettings = new UserSettings(null, null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }
    }
}