using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ExtractMethodTests : VbeTestBase
    {
        [TestMethod]
        public void ExtractMethod_PrivateFunction()
        {
            const string inputCode = @"
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub";

            const string expectedCode = @"
Private Sub Foo()
    x = Bar()
End Sub

Private Function Bar() As Integer
    Dim x As Integer
    x = 1 + 2
    Bar = x
End Function

";
            
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(module.Parent), new Selection(4, 1, 4, 20));
            var model = new ExtractMethodModel(vbe.Object, parser.State.AllDeclarations, qualifiedSelection);
            model.Method.Accessibility = Accessibility.Private;
            model.Method.MethodName = "Bar";
            model.Method.ReturnValue = new ExtractedParameter("Integer", ExtractedParameter.PassedBy.ByVal, "x");
            model.Method.Parameters = new List<ExtractedParameter>();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ExtractMethodRefactoring(vbe.Object, factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        private static Mock<IRefactoringPresenterFactory<IExtractMethodPresenter>> SetupFactory(ExtractMethodModel model)
        {
            var presenter = new Mock<IExtractMethodPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IExtractMethodPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }
    }
}
