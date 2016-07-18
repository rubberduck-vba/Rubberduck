using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ExtractInterfaceTests
    {
        [TestMethod]
        public void ExtractInterfaceRefactoring_ImplementProc()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Implements ITestModule1

Private Sub ITestModule1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string expectedInterfaceCode =
@"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, vbext_ComponentType.vbext_ct_ClassModule, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object, new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
            foreach (var member in model.Members)
            {
                member.IsSelected = true;
            }

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, parser.State, null, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents.Item(1).CodeModule.Lines());
            Assert.AreEqual(expectedCode, project.VBComponents.Item(0).CodeModule.Lines());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_ImplementProcAndFuncAndPropGetSetLet()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Implements ITestModule1

Private Sub ITestModule1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function ITestModule1_Fizz(ByRef b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function

Private Property Get ITestModule1_Buzz() As Variant
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let ITestModule1_Buzz(ByRef value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Set ITestModule1_Buzz(ByRef value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            const string expectedInterfaceCode =
@"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(ByRef b As Variant) As Variant
End Function

Public Property Get Buzz() As Variant
End Property

Public Property Let Buzz(ByRef value As Variant)
End Property

Public Property Set Buzz(ByRef value As Variant)
End Property

";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, vbext_ComponentType.vbext_ct_ClassModule, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object, new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
            foreach (var member in model.Members)
            {
                member.IsSelected = true;
            }

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, parser.State, null, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents.Item(1).CodeModule.Lines());
            Assert.AreEqual(expectedCode, project.VBComponents.Item(0).CodeModule.Lines());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_ImplementProcAndFunc_IgnoreProperties()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Implements ITestModule1

Private Sub ITestModule1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function ITestModule1_Fizz(ByRef b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function


Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            const string expectedInterfaceCode =
@"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(ByRef b As Variant) As Variant
End Function

";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, vbext_ComponentType.vbext_ct_ClassModule, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object, new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
            foreach (var member in model.Members)
            {
                if (!member.FullMemberSignature.Contains("Property"))
                {
                    member.IsSelected = true;
                }
            }

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, parser.State, null, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents.Item(1).CodeModule.Lines());
            Assert.AreEqual(expectedCode, project.VBComponents.Item(0).CodeModule.Lines());
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IExtractInterfacePresenter>> SetupFactory(ExtractInterfaceModel model)
        {
            var presenter = new Mock<IExtractInterfacePresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IExtractInterfacePresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}