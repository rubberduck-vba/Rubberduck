using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
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

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub ITestModule1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
@"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);
            foreach (var member in model.Members)
            {
                member.IsSelected = true;
            }

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, null, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_ImplementProcAndFuncAndPropGetSetLet()
        {
            //Input
            const string inputCode = @"
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

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode = @"
Implements ITestModule1

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property

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
";

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);
            foreach (var member in model.Members)
            {
                member.IsSelected = true;
            }

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, null, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property

Private Sub ITestModule1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function ITestModule1_Fizz(ByRef b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function
";

            const string expectedInterfaceCode =
@"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(ByRef b As Variant) As Variant
End Function

";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);
            foreach (var member in model.Members)
            {
                if (!member.FullMemberSignature.Contains("Property"))
                {
                    member.IsSelected = true;
                }
            }

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, null, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_IgnoresField()
        {
            //Input
            const string inputCode =
@"Public Fizz As Boolean";

            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);
            Assert.AreEqual(0, model.Members.Count());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_NullPresenter_NoChanges()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);

            //SetupFactory
            var factory = SetupFactory(model);
            factory.Setup(f => f.Create()).Returns(value: null);

            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, null, factory.Object);
            refactoring.Refactor();

            Assert.AreEqual(1, vbe.Object.ActiveVBProject.VBComponents.Count());
            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_NullModel_NoChanges()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);

            var presenter = new Mock<IExtractInterfacePresenter>();
            presenter.Setup(p => p.Show()).Returns(value: null);

            //SetupFactory
            var factory = SetupFactory(model);
            factory.Setup(f => f.Create()).Returns(presenter.Object);

            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, null, factory.Object);
            refactoring.Refactor();

            Assert.AreEqual(1, vbe.Object.ActiveVBProject.VBComponents.Count());
            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_PassTargetIn()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Implements ITestModule1

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub ITestModule1_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
@"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(state, qualifiedSelection);
            model.Members.ElementAt(0).IsSelected = true;

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, null, factory.Object);
            refactoring.Refactor(state.AllUserDeclarations.Single(s => s.DeclarationType == DeclarationType.ClassModule));

            Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void Presenter_Reject_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ExtractInterfaceModel(state, qualifiedSelection);
            model.Members.ElementAt(0).IsSelected = true;

            var view = new Mock<IRefactoringDialog<ExtractInterfaceViewModel>>();
            view.Setup(v => v.ViewModel).Returns(new ExtractInterfaceViewModel());
            view.Setup(v => v.DialogResult).Returns(DialogResult.Cancel);

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, state, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_NullTarget_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ExtractInterfaceModel(state, qualifiedSelection);

            var view = new Mock<IRefactoringDialog<ExtractInterfaceViewModel>>();
            view.SetupGet(v => v.ViewModel).Returns(new ExtractInterfaceViewModel());
            var presenter = new ExtractInterfacePresenter(view.Object, model);

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Factory_NoMembersInTarget_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, state, null);

            Assert.AreEqual(null, factory.Create());
        }

        [TestMethod]
        public void Factory_NullSelectionNullReturnsNullPresenter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, state, null);

            Assert.AreEqual(null, factory.Create());
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