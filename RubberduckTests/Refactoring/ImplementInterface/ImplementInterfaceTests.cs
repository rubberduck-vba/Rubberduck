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
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ImplementInterfaceTests : RefactoringTestBase
    {
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

            //Expectation
            const string expectedCode =
                @"Implements Class1

Private Sub Class1_Foo()
    Err.Raise 5 'TODO implement interface member
End Sub
";

            var actualCode = RefactoredCode(
                "Class2", 
                selection, 
                null, 
                false,
                ("Class1", inputCode1, ComponentType.ClassModule), 
                ("Class2", inputCode2, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Class2"]);
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

            const string expectedCode = @"Implements IInterface

Private Sub IInterface_DoSomething()
    Err.Raise 5 'TODO implement interface member
End Sub
";

            var actualCode = RefactoredCode(
                "Sheet1",
                selection,
                null,
                false,
                ("IInterface", interfaceCode, ComponentType.ClassModule),
                ("Sheet1", initialCode, ComponentType.Document));
            Assert.AreEqual(expectedCode, actualCode["Sheet1"]);
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

            const string expectedCode = @"Implements IInterface

Private Sub IInterface_DoSomething()
    Err.Raise 5 'TODO implement interface member
End Sub
";

            var actualCode = RefactoredCode(
                "Form1",
                selection,
                null,
                false,
                ("IInterface", interfaceCode, ComponentType.ClassModule),
                ("Form1", initialCode, ComponentType.UserForm));
            Assert.AreEqual(expectedCode, actualCode["Form1"]);
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

            //Expectation
            const string expectedCode =
                @"Implements Class1
Implements Class3

Private Sub Class1_Foo()
    Err.Raise 5 'TODO implement interface member
End Sub
";

            var actualCode = RefactoredCode(
                "Class2",
                selection,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Class2"]);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state,
            ISelectionService selectionService)
        {
            var addImplementationsBaseRefactoring = new AddInterfaceImplementationsRefactoringAction(rewritingManager, new CodeBuilder());
            var baseRefactoring = new ImplementInterfaceRefactoringAction(addImplementationsBaseRefactoring, rewritingManager);
            return new ImplementInterfaceRefactoring(baseRefactoring, state, selectionService);
        }
    }
}

