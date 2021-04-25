using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.Exceptions.ImplementInterface;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ImplementInterfaceTests : RefactoringTestBase
    {
        private static readonly string errRaise = "Err.Raise 5";
        private static readonly string todoMsg = Rubberduck.Resources.Refactorings.Refactorings.ImplementInterface_TODO;

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public override void TargetNull_Throws()
        {
            var testVbe = TestVbe(string.Empty, out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(testVbe);
            using (state)
            {
                var refactoring = TestRefactoring(rewritingManager, state);
                Assert.Throws<NotSupportedException>(() => refactoring.Refactor((Declaration)null));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void DoesNotSupportCallingWithADeclaration()
        {
            var testVbe = TestVbe(("testClass", string.Empty, ComponentType.ClassModule));
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(testVbe);
            using (state)
            {
                var target = state.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule).Single();
                var refactoring = TestRefactoring(rewritingManager, state);
                Assert.Throws<NotSupportedException>(() => refactoring.Refactor(target));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementInterface_SelectionOnImplementsStatement()
        {
            //Input
            const string inputCode1 =
                @"Public Sub Foo()
End Sub";

            const string inputCode2 =
                @"Implements Class1";

            var selection = Selection.Home;

            var actualCode = RefactoredCode(
                "Class2", 
                selection, 
                null, 
                false,
                ("Class1", inputCode1, ComponentType.ClassModule), 
                ("Class2", inputCode2, ComponentType.ClassModule));

            StringAssert.Contains("Implements Class1", actualCode["Class2"]);
            StringAssert.Contains("Private Sub Class1_Foo()", actualCode["Class2"]);
            StringAssert.Contains(errRaise, actualCode["Class2"]);
            StringAssert.Contains(todoMsg, actualCode["Class2"]);
            StringAssert.Contains("End Sub", actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void DoesNotImplementInterface_SelectionNotOnImplementsStatement()
        {
            //Input
            const string inputCode1 =
                @"Public Sub Foo()
End Sub";

            const string inputCode2 =
                @"Implements Class1
   
";
            var selection = new Selection(2, 2);

            //Expectation
            const string expectedCode =
                @"Implements Class1
   
";

            var actualCode = RefactoredCode(
                "Class2", 
                selection, 
                typeof(NoImplementsStatementSelectedException), 
                false, 
                ("Class1", inputCode1, ComponentType.ClassModule), 
                ("Class2", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementsInterfaceInDocumentModule()
        {
            const string interfaceCode = @"Option Explicit
Public Sub DoSomething()
End Sub
";
            const string initialCode = @"Implements IInterface";

            var selection = Selection.Home;

            var actualCode = RefactoredCode(
                "Sheet1",
                selection,
                null,
                false,
                ("IInterface", interfaceCode, ComponentType.ClassModule),
                ("Sheet1", initialCode, ComponentType.Document));

            StringAssert.Contains("Implements IInterface", actualCode["Sheet1"]);
            StringAssert.Contains("Private Sub IInterface_DoSomething()", actualCode["Sheet1"]);
            StringAssert.Contains(errRaise, actualCode["Sheet1"]);
            StringAssert.Contains(todoMsg, actualCode["Sheet1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementsInterfaceInUserFormModule()
        {
            const string interfaceCode = @"Option Explicit
Public Sub DoSomething()
End Sub
";
            const string initialCode = @"Implements IInterface";

            var selection = Selection.Home;

            var actualCode = RefactoredCode(
                "Form1",
                selection,
                null,
                false,
                ("IInterface", interfaceCode, ComponentType.ClassModule),
                ("Form1", initialCode, ComponentType.UserForm));

            StringAssert.Contains("Implements IInterface", actualCode["Form1"]);
            StringAssert.Contains("Private Sub IInterface_DoSomething()", actualCode["Form1"]);
            StringAssert.Contains(errRaise, actualCode["Form1"]);
            StringAssert.Contains(todoMsg, actualCode["Form1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Implement Interface")]
        public void ImplementsCorrectInterface_MultipleInterfaces()
        {
            //Input
            const string inputCode1 =
                @"Public Sub Foo()
End Sub";

            const string inputCode2 =
                @"Implements Class1
Implements Class3";

            var selection = new Selection(1, 1);

            const string inputCode3 =
                @"Public Sub Foo()
End Sub";

            var actualCode = RefactoredCode(
                "Class2",
                selection,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule));

            StringAssert.Contains("Implements Class3", actualCode["Class2"]);
            StringAssert.Contains(errRaise, actualCode["Class2"]);
            StringAssert.Contains(todoMsg, actualCode["Class2"]);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state,
            ISelectionService selectionService)
        {
            var addImplementationsBaseRefactoring = new AddInterfaceImplementationsRefactoringAction(rewritingManager, CreateCodeBuilder());
            var baseRefactoring = new ImplementInterfaceRefactoringAction(addImplementationsBaseRefactoring, rewritingManager);
            return new ImplementInterfaceRefactoring(baseRefactoring, state, selectionService);
        }

        private static ICodeBuilder CreateCodeBuilder()
            => new CodeBuilder(new Indenter(null, CreateIndenterSettings));

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }
    }
}

