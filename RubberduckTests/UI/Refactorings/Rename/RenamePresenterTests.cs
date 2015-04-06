using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings.Rename;
using parsing = Rubberduck.Parsing;

namespace RubberduckTests.UI.Refactorings.Rename
{
    [TestClass]
    public class RenamePresenterTests
    {
        private static Mock<VBE> _vbe;
        private static Mock<VBProject> _vbProject;
        private static Declarations _declarations;
        private static parsing.QualifiedModuleName _module;
        private static Mock<IRenameView> _view;
        private List<Declaration> _listDeclarations;

        [TestInitialize]
        public void TestInitialization()
        {
            _vbe = new Mock<VBE>();
            _vbProject = new Mock<VBProject>();
            _declarations = new Declarations();
            _module = new parsing.QualifiedModuleName();
            _view = new Mock<IRenameView>();
        }

        /// <summary>Common method for adding declaration with some default values</summary>
        private void AddDeclarationItem(IMock<ParserRuleContext> context,
            parsing.Selection selection,
            parsing.QualifiedMemberName? qualifiedName = null,
            DeclarationType declarationType = DeclarationType.Project,
            string identifierName = "identifierName")
        {
            Declaration decItem = null;
            var qualName = qualifiedName ?? new parsing.QualifiedMemberName(_module, "fakeModule");

            decItem = new Declaration(qualName,
                accessibility: Accessibility.Public,
                declarationType: declarationType,
                context: context.Object,
                selection: selection,
                parentScope: "parentScope",
                identifierName: identifierName,
                asTypeName: "asTypeName",
                isSelfAssigned: false);

            _declarations.Add(decItem);
            if (_listDeclarations == null) _listDeclarations = new List<Declaration>();
            _listDeclarations.Add(decItem);
        }

        /// <summary>Common method for adding a reference to given declaration item</summary>
        private static void AddReference(Declaration itemToAdd, IdentifierReference reference)
        {
            var declaration = _declarations.Items.ToList().FirstOrDefault(x => x.Equals(itemToAdd));
            if (declaration == null) return;

            declaration.AddReference(reference);
        }

        [TestMethod]
        public void ConstructorWorks_IsNotNull()
        {
            // arange
            var symbolSelection = new parsing.Selection(1, 1, 2, 2);
            var qualifiedSelection = new parsing.QualifiedSelection(_module, symbolSelection);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);

            //assert
            Assert.IsNotNull(presenter, "Successfully initialized");
        }

        [TestMethod]
        public void NoTargetFound()
        {
            // arange
            var symbolSelection = new parsing.Selection(1, 1, 2, 2);
            var qualifiedSelection = new parsing.QualifiedSelection(_module, symbolSelection);

            var context = new Mock<ParserRuleContext>();
            AddDeclarationItem(context, symbolSelection);
            _view.Setup(form => form.ShowDialog()).Returns(DialogResult.Cancel);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);
            presenter.Show();

            //assert
            Assert.IsNull(_view.Object.Target, "No Target was found");
        }

        [TestMethod]
        public void AcquireTarget_ProcedureRenaming_TargetIsNotNull()
        {
            // arange
            var symbolSelection = new parsing.Selection(1, 1, 2, 4);
            var qualifiedSelection = new parsing.QualifiedSelection(_module, symbolSelection);

            // just for passing null reference exception
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(1);
            context.SetupGet(c => c.Stop.Line).Returns(2);
            context.SetupGet(c => c.Stop.Text).Returns("Four");

            // setting a declaration item as a procedure that will be renamed
            const string identifierName = "AProcedure";
            AddDeclarationItem(context, symbolSelection, null, DeclarationType.Procedure, identifierName);

            // allow Moq to set the Target property
            _view.Setup(view => view.ShowDialog()).Returns(DialogResult.Cancel);
            _view.SetupProperty(view => view.Target);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);
            presenter.Show();

            //assert
            Assert.IsNotNull(_view.Object.Target, "A target was found");
            Assert.AreEqual(identifierName, _view.Object.Target.IdentifierName);
        }

        [TestMethod]
        public void AcquireTarget_ModuleRenaming_TargetIsNotNull()
        {
            // arange
            var symbolSelection = new parsing.Selection(1, 1, 2, 4);
            var qualifiedSelection = new parsing.QualifiedSelection(_module, symbolSelection);

            // just for passing null reference exception
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(1);
            context.SetupGet(c => c.Stop.Line).Returns(2);
            context.SetupGet(c => c.Stop.Text).Returns("Four");

            // setting a declaration item as a module that will be renamed
            const string identifierName = "FakeModule";
            AddDeclarationItem(context, symbolSelection, null, DeclarationType.Module, identifierName);

            // allow Moq to set the Target property
            _view.Setup(view => view.ShowDialog()).Returns(DialogResult.Cancel);
            _view.SetupProperty(view => view.Target);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);
            presenter.Show();

            //assert
            Assert.IsNotNull(_view.Object.Target, "A target was found");
            Assert.AreEqual(identifierName, _view.Object.Target.IdentifierName);
        }

        [TestMethod]
        public void AcquireTarget_MethodRenamingAtSameComponent_CorrectTargetChosen()
        {
            // arange
            var symbolSelection = new parsing.Selection(8, 1, 8, 16);
            var selectedComponent = new parsing.QualifiedModuleName("TestProject", "TestModule", _vbProject.Object, 1);
            var qualifiedSelection = new parsing.QualifiedSelection(selectedComponent, symbolSelection);

            // just for passing null reference exception            
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(1);
            context.SetupGet(c => c.Stop.Line).Returns(2);
            context.SetupGet(c => c.Stop.Text).Returns("Four");

            // simulate all the components and symbols   
            var member = new parsing.QualifiedMemberName(selectedComponent, "fakeModule");
            const string identifierName = "Foo";
            AddDeclarationItem(context, symbolSelection, member, DeclarationType.Procedure, identifierName);
            AddDeclarationItem(context, new parsing.Selection(1, 1, 1, 16), member, DeclarationType.Procedure);
            AddDeclarationItem(context, new parsing.Selection(1, 1, 1, 1), member, DeclarationType.Module);

            // allow Moq to set the Target property
            _view.Setup(view => view.ShowDialog()).Returns(DialogResult.Cancel);
            _view.SetupProperty(view => view.Target);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);
            presenter.Show();

            //assert
            var retVal = _view.Object.Target;
            Assert.AreEqual(symbolSelection, retVal.Selection, "Returns only the declaration on the desired selection");
            Assert.AreEqual(identifierName, retVal.IdentifierName);
        }

        [TestMethod]
        public void AcquireTarget_MethodRenamingMoreComponents_CorrectTargetChosen()
        {
            // arange
            // initial selection
            var symbolSelection = new parsing.Selection(4, 5, 4, 8);
            var selectedComponent = new parsing.QualifiedModuleName("TestProject", "Module1", _vbProject.Object, 1);
            var qualifiedSelection = new parsing.QualifiedSelection(selectedComponent, symbolSelection);

            // just for passing null reference exception            
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(-1);
            context.SetupGet(c => c.Stop.Line).Returns(-1);
            context.SetupGet(c => c.Stop.Text).Returns("Fake");

            // simulate all the components and symbols   
            IdentifierReference reference;
            var differentComponent = new parsing.QualifiedModuleName("TestProject", "Module2", _vbProject.Object, 1);
            var differentMember = new parsing.QualifiedMemberName(differentComponent, "Module2");
            AddDeclarationItem(context, new parsing.Selection(4, 9, 4, 16), differentMember, DeclarationType.Variable,"FooTest");

            // add references to the Foo declaration item to simulate prod usage
            AddDeclarationItem(context, new parsing.Selection(3, 5, 3, 8), differentMember, DeclarationType.Procedure, "Foo");
            var declarationItem = _listDeclarations[_listDeclarations.Count - 1];
            reference = new IdentifierReference(selectedComponent, "Foo", new parsing.Selection(7, 5, 7, 11), false,context.Object, declarationItem);
            AddReference(declarationItem, reference);
            reference = new IdentifierReference(selectedComponent, "Foo", symbolSelection, false, context.Object,declarationItem);
            AddReference(declarationItem, reference);

            AddDeclarationItem(context, new parsing.Selection(1, 1, 1, 1), differentMember, DeclarationType.Module, "Module2");
            var member = new parsing.QualifiedMemberName(selectedComponent, "fakeModule");
            AddDeclarationItem(context, new parsing.Selection(7, 5, 7, 11), member, DeclarationType.Procedure, "RunFoo");
            AddDeclarationItem(context, new parsing.Selection(3, 5, 3, 9), member, DeclarationType.Procedure, "Main");
            AddDeclarationItem(context, new parsing.Selection(1, 1, 1, 1), member, DeclarationType.Module, "Module1");

            _view.Setup(view => view.ShowDialog()).Returns(DialogResult.Cancel);
            _view.SetupProperty(view => view.Target);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);
            presenter.Show();

            //assert
            var retVal = _view.Object.Target;
            Assert.AreEqual("Foo", retVal.IdentifierName, "Selected the correct symbol name");
            Assert.AreEqual(declarationItem.References.Count(), retVal.References.Count());
        }
    }
}