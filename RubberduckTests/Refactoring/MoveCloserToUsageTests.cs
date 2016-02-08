using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
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
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()

    Dim bar As Boolean
    bar = True
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(2, 1, 2, 1);

            //Expectation
            const string expectedCode =
@"Private bar As Integer
Private bay As Date

Private Sub Foo()

    Dim bat As Boolean
    bat = True
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(6, 6, 6, 6);

            //Expectation
            const string expectedCode =
@"Private           bat As Boolean,          bay As Date

Private Sub Foo()

    Dim bar As Integer
    bar = 3
End Sub";   // note: VBE will remove extra spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(6, 6, 6, 6);

            //Expectation
            const string expectedCode =
@"Private bar As Integer,          bay As Date

Private Sub Foo()

    Dim bat As Boolean
    bat = True
End Sub";   // note: VBE will remove extra spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(6, 6, 6, 6);

            //Expectation
            const string expectedCode =
@"Private bar As Integer,          bat As Boolean

Private Sub Foo()

    Dim bay As Date
    bay = #1/13/2004#
End Sub";   // note: VBE will remove extra spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(2, 16, 2, 16);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim         bat As Boolean,        bay As Date

    bat = True

    Dim bar As Integer
    bar = 3
End Sub";   // note: VBE will remove extra spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(3, 16, 3, 16);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim bar As Integer,        bay As Date

    bar = 1

    Dim bat As Boolean
    bat = True
End Sub";   // note: VBE will remove extra spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(4, 16, 4, 16);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim bar As Integer,        bat As Boolean

    bar = 4

    Dim bay As Date
    bay = #1/13/2004#
End Sub";   // note: VBE will remove extra spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void MoveCloserToUsageRefactoring_NoReferences()
        {
            //Input
            const string inputCode =
@"Private bar As Boolean
Private Sub Foo()
End Sub";
            var selection = new Selection(1, 1, 1, 1);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, module.Lines());
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
            var selection = new Selection(1, 1, 1, 1);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, module.Lines());
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
            var selection = new Selection(1, 1, 1, 1);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 1, 1, 1);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 1, 1, 1);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void MoveCloserToUsageRefactoring_ReferenceIsSeparatedWithColon()
        {
            //Input
            const string inputCode =
@"Private bar As Boolean
Private Sub Foo(): Baz True, True, bar: End Sub
Private Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean): End Sub";

            var selection = new Selection(1, 1, 1, 1);

            // Yeah, this code is a mess.  That is why we got the SmartIndenter
            const string expectedCode =
@"Private Sub Foo()
    Dim bar As Boolean

Baz True, True, bar
End Sub
Private Sub Baz(ByVal bat As Boolean, ByVal bas As Boolean, ByVal bac As Boolean): End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);

            //Assert
            try
            {
                refactoring.Refactor(parser.State.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Invalid Argument. DeclarationType must be 'Variable'\r\nParameter name: target", e.Message);
                Assert.AreEqual(inputCode, module.Lines());
                messageBox.Verify(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>()), Times.Once);
                return;
            }

            Assert.Fail();
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
            var selection = new Selection(2, 15, 2, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new MoveCloserToUsageRefactoring(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                    It.IsAny<MessageBoxIcon>()), Times.Once);

            Assert.AreEqual(inputCode, module.Lines());
        }
    }
}
