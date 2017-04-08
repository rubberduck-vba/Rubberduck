using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class MoveCloserToUsageTests
    {
        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            var state = MockParser.CreateAndParse(vbe.Object);

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

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void MoveCloserToUsageRefactoring_NoReferences()
        {
            //Input
            const string inputCode =
@"Private bar As Boolean
Private Sub Foo()
End Sub";
            var selection = new Selection(1, 1);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
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
@"Public Sub Test(): Dim foo As Long
SomeSub someParam:=foo: End Sub

Public Sub SomeSub(ByVal someParam As Long)
    Debug.Print someParam
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_PassInTarget_Nonvariable()
        {
            //Input
            const string inputCode =
@"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(state.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));
            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());

            messageBox.Verify(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_InvalidSelection()
        {
            //Input
            const string inputCode =
@"Private bar As Boolean
Private Sub Foo()
    bar = True
End Sub";
            var selection = new Selection(2, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new MoveCloserToUsageRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                    It.IsAny<MessageBoxIcon>()), Times.Once);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());
        }
    }
}
