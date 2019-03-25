using System.Collections.ObjectModel;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractInterfaceTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, qualifiedSelection);
                foreach (var interfaceMember in model.Members)
                {
                    interfaceMember.IsSelected = true;
                }

                var refactoring = TestRefactoring(vbe.Object, rewritingManager, state, model);
                refactoring.Refactor(qualifiedSelection);
                var actualCode = component.CodeModule.Content();

                Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, qualifiedSelection);
                foreach (var interfaceMember in model.Members)
                {
                    interfaceMember.IsSelected = true;
                }

                var refactoring = TestRefactoring(vbe.Object, rewritingManager, state, model);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, qualifiedSelection);
                foreach (var interfaceMember in model.Members.Where(member => !member.FullMemberSignature.Contains("Property")))
                {
                    interfaceMember.IsSelected = true;
                }

                var refactoring = TestRefactoring(vbe.Object, rewritingManager, state, model);
                refactoring.Refactor(qualifiedSelection);

                Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_IgnoresField()
        {
            //Input
            const string inputCode =
                @"Public Fizz As Boolean";

            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, qualifiedSelection);
                Assert.AreEqual(0, model.Members.Count());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NullPresenter_NoChanges()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, qualifiedSelection);

                //SetupFactory
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IExtractInterfacePresenter, ExtractInterfaceModel>(It.IsAny<ExtractInterfaceModel>())).Returns(value: null);

                var refactoring = TestRefactoring(vbe.Object, rewritingManager, state, factory.Object);
                refactoring.Refactor();

                Assert.AreEqual(1, vbe.Object.ActiveVBProject.VBComponents.Count());
                Assert.AreEqual(inputCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NullModel_NoChanges()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var refactoring = TestRefactoring(vbe.Object, rewritingManager, state, model: null);
                refactoring.Refactor();

                Assert.AreEqual(1, vbe.Object.ActiveVBProject.VBComponents.Count());
                Assert.AreEqual(inputCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, qualifiedSelection);
                model.Members.ElementAt(0).IsSelected = true;
                model.Members = new ObservableCollection<InterfaceMember>(new[] {model.Members.ElementAt(0)}.ToList());
                
                var refactoring = TestRefactoring(vbe.Object, rewritingManager, state, model);
                refactoring.Refactor(state.AllUserDeclarations.Single(s => s.DeclarationType == DeclarationType.ClassModule));

                Assert.AreEqual(expectedInterfaceCode, component.Collection[1].CodeModule.Content());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        #region setup
        private static IRefactoring TestRefactoring(IVBE vbe, IRewritingManager rewritingManager, RubberduckParserState state, ExtractInterfaceModel model, IMessageBox msgBox = null)
        {
            var factory = SetupFactory(model);
            return TestRefactoring(vbe, rewritingManager, state, factory.Object, msgBox);
        }

        private static IRefactoring TestRefactoring(IVBE vbe, IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, IMessageBox msgBox = null)
        {
            var selectionService = MockedSelectionService(vbe.GetActiveSelection());
            if (msgBox == null)
            {
                msgBox = new Mock<IMessageBox>().Object;
            }
            return new ExtractInterfaceRefactoring(state, state, msgBox, factory, rewritingManager, selectionService);
        }

        private static ISelectionService MockedSelectionService(QualifiedSelection? initialSelection)
        {
            QualifiedSelection? activeSelection = initialSelection;
            var selectionServiceMock = new Mock<ISelectionService>();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            selectionServiceMock.Setup(m => m.TrySetActiveSelection(It.IsAny<QualifiedSelection>()))
                .Returns(() => true).Callback((QualifiedSelection selection) => activeSelection = selection);
            return selectionServiceMock.Object;
        }

        private static Mock<IRefactoringPresenterFactory> SetupFactory(ExtractInterfaceModel model)
        {
            var presenter = new Mock<IExtractInterfacePresenter>();

            var factory = new Mock<IRefactoringPresenterFactory>();
            factory.Setup(f => f.Create<IExtractInterfacePresenter, ExtractInterfaceModel>(It.IsAny<ExtractInterfaceModel>()))
                .Callback(() => presenter.Setup(p => p.Show()).Returns(model))
                .Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}