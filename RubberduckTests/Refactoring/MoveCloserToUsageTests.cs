using System;
using System.Linq;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class MoveCloserToUsageTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Field()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_LineNumbers()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
100 bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
Dim bar As Boolean
100 bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Field_MultipleLines()
        {
            //Input
            const string inputCode =
                @"Private _
bar _
As _
Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_FieldInOtherClass()
        {
            //Input
            const string inputCode1 =
                @"Public bar As Boolean";

            const string inputCode2 =
                @"Private Sub Foo()
Module1.bar = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode1 =
                @"";

            const string expectedCode2 =
                @"Private Sub Foo()
Dim bar As Boolean
bar = True
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);
                var module1 = project.Object.VBComponents[0];
                var module2 = project.Object.VBComponents[1];

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter1 = state.GetRewriter(module1);
                Assert.AreEqual(expectedCode1, rewriter1.GetText());

                var rewriter2 = state.GetRewriter(module2);
                Assert.AreEqual(expectedCode2, rewriter2.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Variable()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    Dim bat As Integer
    bar = True
End Sub";
            var selection = new Selection(4, 6, 4, 8);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bat As Integer
    Dim bar As Boolean
    bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_VariableWithLineNumbers()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
1   Dim bar As Boolean
2   Dim bat As Integer
3   bar = True
End Sub";
            var selection = new Selection(4, 6, 4, 8);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
1   
2   Dim bat As Integer
Dim bar As Boolean
3   bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_Variable_MultipleLines()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim _
    bar _
    As _
    Boolean
    Dim bat As Integer
    bar = True
End Sub";
            var selection = new Selection(4, 6, 4, 8);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bat As Integer
    Dim bar As Boolean
    bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFields_MoveSecond()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer
Private bat As Boolean
Private bay As Date

Private Sub Foo()
    bat = True
End Sub";
            var selection = new Selection(2, 1);

            //Expectation
            const string expectedCode =
                @"Private bar As Integer
Private bay As Date

Private Sub Foo()
    Dim bat As Boolean
    bat = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFieldsOneStatement_MoveFirst()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer, _
          bat As Boolean, _
          bay As Date

Private Sub Foo()
    bar = 3
End Sub";
            var selection = new Selection(6, 6);

            //Expectation
            const string expectedCode =
                @"Private bat As Boolean, _
          bay As Date

Private Sub Foo()
    Dim bar As Integer
    bar = 3
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFieldsOneStatement_MoveSecond()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer, _
          bat As Boolean, _
          bay As Date

Private Sub Foo()
    bat = True
End Sub";
            var selection = new Selection(6, 6);

            //Expectation
            const string expectedCode =
                @"Private bar As Integer, _
          bay As Date

Private Sub Foo()
    Dim bat As Boolean
    bat = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleFieldsOneStatement_MoveLast()
        {
            //Input
            const string inputCode =
                @"Private bar As Integer, _
          bat As Boolean, _
          bay As Date

Private Sub Foo()
    bay = #1/13/2004#
End Sub";
            var selection = new Selection(6, 6);

            //Expectation
            const string expectedCode =
                @"Private bar As Integer, _
          bat As Boolean

Private Sub Foo()
    Dim bay As Date
    bay = #1/13/2004#
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleVariablesOneStatement_MoveFirst()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean, _
        bay As Date

    bat = True
    bar = 3
End Sub";
            var selection = new Selection(2, 16);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bat As Boolean, _
        bay As Date

    bat = True
    Dim bar As Integer
    bar = 3
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleVariablesOneStatement_MoveSecond()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean, _
        bay As Date

    bar = 1
    bat = True
End Sub";
            var selection = new Selection(3, 16);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bay As Date

    bar = 1
    Dim bat As Boolean
    bat = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultipleVariablesOneStatement_MoveLast()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean, _
        bay As Date

    bar = 4
    bay = #1/13/2004#
End Sub";
            var selection = new Selection(4, 16);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Integer, _
        bat As Boolean

    bar = 4
    Dim bay As Date
    bay = #1/13/2004#
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_NoReferences()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
End Sub";
            var selection = new Selection(1, 1);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var messageBox = new Mock<IMessageBox>();

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
                refactoring.Refactor(qualifiedSelection);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(inputCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferencedInMultipleProcedures()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub
Private Sub Bar()
    bar = True
End Sub";
            var selection = new Selection(1, 1);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var messageBox = new Mock<IMessageBox>();

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
                refactoring.Refactor(qualifiedSelection);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once());
                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(inputCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsNotBeginningOfStatement_Assignment()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo(ByRef bat As Boolean)
    bat = bar
End Sub";

            const string expectedCode =
                @"Private Sub Foo(ByRef bat As Boolean)
    Dim bar As Boolean
    bat = bar
End Sub";
            var selection = new Selection(1, 1);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsNotBeginningOfStatement_PassAsParam()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    Baz bar
End Sub
Sub Baz(ByVal bat As Boolean)
End Sub";

            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    Baz bar
End Sub
Sub Baz(ByVal bat As Boolean)
End Sub";
            var selection = new Selection(1, 1);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsNotBeginningOfStatement_PassAsParam_ReferenceIsNotFirstLine()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    Baz True, _
        True, _
        bar
End Sub
Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean)
End Sub";

            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As Boolean
    Baz True, _
        True, _
        bar
End Sub
Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean)
End Sub";
            var selection = new Selection(1, 1);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_ReferenceIsSeparatedWithColon()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo(): Baz True, True, bar: End Sub
Private Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean): End Sub";

            var selection = new Selection(1, 1);

            // Yeah, this code is a mess.  That is why we got the SmartIndenter
            const string expectedCode =
                @"Private Sub Foo(): Dim bar As Boolean
Baz True, True, bar: End Sub
Private Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean): End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_WorksWithNamedParameters()
        {
            //Input
            const string inputCode =
                @"
Private foo As Long

Public Sub Test()
    SomeSub someParam:=foo
End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var selection = new Selection(2, 1);
            const string expectedCode =
                @"

Public Sub Test()
    Dim foo As Long
    SomeSub someParam:=foo
End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_WorksWithNamedParametersAndStatementSeparaters()
        {
            //Input
            const string inputCode =
                @"Private foo As Long

Public Sub Test(): SomeSub someParam:=foo: End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var selection = new Selection(1, 1);
            const string expectedCode =
                @"
Public Sub Test(): Dim foo As Long
SomeSub someParam:=foo: End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void IntroduceFieldRefactoring_PassInTarget_Nonvariable()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var messageBox = new Mock<IMessageBox>();

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
                refactoring.Refactor(state.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));
                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(inputCode, rewriter.GetText());

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void IntroduceFieldRefactoring_InvalidSelection()
        {
            //Input
            const string inputCode =
                @"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(2, 15);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var messageBox = new Mock<IMessageBox>();

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
                refactoring.Refactor(qualifiedSelection);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(inputCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Move Closer")]
        [Category("Refactorings")]
        public void MoveCloser_RespectsMemberAccess_ContextOwners()
        {
            const string inputCode =
@"
Public Sub foo()
  Dim count As Long
  Dim report As Worksheet
  Set report = ThisWorkbook.ActiveSheet
  With report
    For count = 1 To 10
      If .Cells(1, count) > count Then
        .Cells(2, count).Value2 = count
      End If
    Next
  End With
End Sub";

            const string expectedCode =
@"
Public Sub foo()
  Dim report As Worksheet
  Set report = ThisWorkbook.ActiveSheet
  With report
    Dim count As Long
    For count = 1 To 10
      If .Cells(1, count) > count Then
        .Cells(2, count).Value2 = count
      End If
    Next
  End With
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var qualifiedSelection = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Where(d => d.IdentifierName == "count")
                    .Single().QualifiedSelection;

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.NotNull(rewriter);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }


        [Test]
        [Category("Move Closer")]
        [Category("Refactorings")]
        public void MoveCloser_RespectsObjectProperties_InUsages()
        {
            string inputClassCode =
@"
Option Explicit

Private _name As Long
Private _myOtherProperty As Long

Public Property Set Name(name As String)
    _name = name
End Property

Public Property Get Name() As String
    Name = _name
End Property

Public Property Set OtherProperty(val As Long)
    _myOtherProperty = val
End Property

Public Property Get OtherProperty() As Long
    OtherProperty = _myOtherProperty
End Property

";
            string inputCode = @"Private foo As Class1


Public Sub Test()
    Debug.Print ""Some statements between""
    Debug.Print ""Declaration and first usage!""
    Set foo = new Class1
    foo.Name = ""FooName""
    foo.OtherProperty = 1626
End Sub";

            var selection = new Selection(1, 1);

            const string expected = @"

Public Sub Test()
    Debug.Print ""Some statements between""
    Debug.Print ""Declaration and first usage!""
    Dim foo As Class1
    Set foo = new Class1
    foo.Name = ""FooName""
    foo.OtherProperty = 1626
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected);
            project.AddComponent("Class1", ComponentType.ClassModule, inputClassCode);
            project.AddComponent("Module1", ComponentType.StandardModule, inputCode);
            builder = builder.AddProject(project.Build());
            var vbe = builder.Build();

            var testComponent = project.MockComponents.Find(mc => mc.Object.Name.Equals("Module1"));
            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(testComponent.Object), selection);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
                refactoring.Refactor(qualifiedSelection);
                var rewriter = state.GetRewriter(testComponent.Object);
                var actual = rewriter.GetText();
                Assert.AreEqual(expected, actual);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_DynamicArray()
        {
            //Input
            const string inputCode =
                @"Private bar() As Boolean
Private Sub Foo()
    ReDim bar(0)
    bar(0) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar() As Boolean
    ReDim bar(0)
    bar(0) = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_FixedArray()
        {
            //Input
            const string inputCode =
                @"Private bar(0) As Boolean
Private Sub Foo()
    bar(0) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar(0) As Boolean
    bar(0) = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_FixedArrayBounded()
        {
            //Input
            const string inputCode =
                @"Private bar(1 To 42) As Boolean
Private Sub Foo()
    bar(1) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar(1 To 42) As Boolean
    bar(1) = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_MultiDimensionalArray()
        {
            //Input
            const string inputCode =
                @"Private bar(1, 1) As Boolean
Private Sub Foo()
    bar(0, 0) = True
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar(1, 1) As Boolean
    bar(0, 0) = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Move Closer")]
        public void MoveCloserToUsageRefactoring_SelfAssigned()
        {
            //Input
            const string inputCode =
                @"Private bar As New Collection
Private Sub Foo()
    bar.Add 42
End Sub";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo()
    Dim bar As New Collection
    bar.Add 42
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
                refactoring.Refactor(qualifiedSelection);

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }
    }
}
