using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
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
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var project = vbe.Object.VBProjects[0];
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

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
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, project.VBComponents[0].CodeModule.Content());
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
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var project = vbe.Object.VBProjects[0];
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

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
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, project.VBComponents[0].CodeModule.Content());
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
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var project = vbe.Object.VBProjects[0];
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

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
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, project.VBComponents[0].CodeModule.Content());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_IgnoresField()
        {
            //Input
            const string inputCode =
@"Public Fizz As Boolean";

            var selection = new Selection(1, 23, 1, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
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

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var project = vbe.Object.VBProjects[0];
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);

            //SetupFactory
            var factory = SetupFactory(model);
            factory.Setup(f => f.Create()).Returns(value: null);

            //Act
            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, parser.State, null, factory.Object);
            refactoring.Refactor();

            //Assert
            Assert.AreEqual(1, project.VBComponents.Count());
            Assert.AreEqual(inputCode, project.VBComponents[0].CodeModule.Content());
        }

        [TestMethod]
        public void ExtractInterfaceRefactoring_NullModel_NoChanges()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var project = vbe.Object.VBProjects[0];
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);

            var presenter = new Mock<IExtractInterfacePresenter>();
            presenter.Setup(p => p.Show()).Returns(value: null);

            //SetupFactory
            var factory = SetupFactory(model);
            factory.Setup(f => f.Create()).Returns(presenter.Object);

            //Act
            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, parser.State, null, factory.Object);
            refactoring.Refactor();

            //Assert
            Assert.AreEqual(1, project.VBComponents.Count());
            Assert.AreEqual(inputCode, project.VBComponents[0].CodeModule.Content());
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
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var project = vbe.Object.VBProjects[0];
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
            model.Members.ElementAt(0).IsSelected = true;
            
            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ExtractInterfaceRefactoring(vbe.Object, parser.State, null, factory.Object);
            refactoring.Refactor(parser.State.AllUserDeclarations.Single(s => s.DeclarationType == DeclarationType.ClassModule));

            //Assert
            Assert.AreEqual(expectedInterfaceCode, project.VBComponents[1].CodeModule.Content());
            Assert.AreEqual(expectedCode, project.VBComponents[0].CodeModule.Content());
        }

        [TestMethod]
        public void Presenter_Reject_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
            model.Members.ElementAt(0).IsSelected = true;

            var view = new Mock<IExtractInterfaceDialog>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, parser.State, view.Object);

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

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);

            var view = new Mock<IExtractInterfaceDialog>();
            var presenter = new ExtractInterfacePresenter(view.Object, model);

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsUpdatedModel()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, selection);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ExtractInterfaceModel(parser.State, qualifiedSelection);
            model.Members.ElementAt(0).IsSelected = true;

            var view = new Mock<IExtractInterfaceDialog>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);
            view.Setup(v => v.InterfaceName).Returns("Class1");

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, parser.State, view.Object);
            var presenter = factory.Create();

            Assert.AreEqual("Class1", presenter.Show().InterfaceName);
        }

        [TestMethod]
        public void Factory_NoMembersInTarget_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("Module1", ComponentType.StandardModule, inputCode, selection);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, parser.State, null);

            Assert.AreEqual(null, factory.Create());
        }

        [TestMethod]
        public void Factory_NullSelectionNullReturnsNullPresenter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("Module1", ComponentType.ClassModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var factory = new ExtractInterfacePresenterFactory(vbe.Object, parser.State, null);

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