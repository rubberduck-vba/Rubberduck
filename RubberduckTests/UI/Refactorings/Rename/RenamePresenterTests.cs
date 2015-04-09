using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings.Rename;
using Parsing = Rubberduck.Parsing;

namespace RubberduckTests.UI.Refactorings.Rename
{
    [TestClass]
    public class RenamePresenterTests
    {
        private static Mock<VBE> _vbe;
        private static Mock<VBProject> _vbProject;
        private static Declarations _declarations;
        private static Parsing.QualifiedModuleName _module;
        private static Mock<IRenameView> _view;
        private List<Declaration> _listDeclarations;

        [TestInitialize]
        public void TestInitialization()
        {
            _vbe = new Mock<VBE>();
            _vbProject = new Mock<VBProject>();
            _declarations = new Declarations();
            _module = new Parsing.QualifiedModuleName();
            _view = new Mock<IRenameView>();
        }

        /// <summary>Common method for adding declaration with some default values</summary>
        private void AddDeclarationItem(IMock<ParserRuleContext> context,
            Parsing.Selection selection,
            Parsing.QualifiedMemberName? qualifiedName = null,
            DeclarationType declarationType = DeclarationType.Project,
            string identifierName = "identifierName")
        {
            Declaration declarationItem = null;
            var qualName = qualifiedName ?? new Parsing.QualifiedMemberName(_module, "fakeModule");

            declarationItem = new Declaration(qualName,
                accessibility: Accessibility.Public,
                declarationType: declarationType,
                context: context.Object,
                selection: selection,
                parentScope: "parentScope",
                identifierName: identifierName,
                asTypeName: "asTypeName",
                isSelfAssigned: false);

            _declarations.Add(declarationItem);
            if (_listDeclarations == null) _listDeclarations = new List<Declaration>();
            _listDeclarations.Add(declarationItem);
        }

        /// <summary>Common method for adding a reference to given declaration item</summary>
        private static void AddReference(Declaration itemToAdd, IdentifierReference reference)
        {
            var declaration = _declarations.Items.ToList().FirstOrDefault(x => x.Equals(itemToAdd));
            if (declaration == null) return;

            declaration.AddReference(reference);
        }

        /// <summary>Common method for creating a VBProject</summary>
        private static Mock<VBComponent> AddComponent(string componentName, CodeModule codeModule)
        {
            Mock<VBComponent> vbComponnet = new Mock<VBComponent>();
            vbComponnet.Setup(x => x.Name).Returns(componentName);
            vbComponnet.Setup(x => x.CodeModule).Returns(codeModule);
            return vbComponnet;
        }

        /// <summary>Common method for creating a VBProject</summary>
        private static Mock<VBProject> AddProject(string projectName, VBComponents vbComponents, vbext_ProjectProtection projectProtection = vbext_ProjectProtection.vbext_pp_none)
        {
            Mock<VBProject> vbProject = new Mock<VBProject>();
            vbProject.SetupGet(x => x.Name).Returns(projectName);
            vbProject.SetupGet(x => x.Protection).Returns(projectProtection);
            vbProject.Setup(x => x.VBComponents).Returns(vbComponents);
            return vbProject;
        }

        [TestMethod]
        public void ConstructorWorks_IsNotNull()
        {
            // arange
            var symbolSelection = new Parsing.Selection(1, 1, 2, 2);
            var qualifiedSelection = new Parsing.QualifiedSelection(_module, symbolSelection);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);

            //assert
            Assert.IsNotNull(presenter, "Successfully initialized");
        }

        [TestMethod]
        public void NoTargetFound()
        {
            // arange
            var symbolSelection = new Parsing.Selection(1, 1, 2, 2);
            var qualifiedSelection = new Parsing.QualifiedSelection(_module, symbolSelection);

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
            var symbolSelection = new Parsing.Selection(1, 1, 2, 4);
            var qualifiedSelection = new Parsing.QualifiedSelection(_module, symbolSelection);

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
            var symbolSelection = new Parsing.Selection(1, 1, 2, 4);
            var qualifiedSelection = new Parsing.QualifiedSelection(_module, symbolSelection);

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
            var symbolSelection = new Parsing.Selection(8, 1, 8, 16);
            var selectedComponent = new Parsing.QualifiedModuleName("TestProject", "TestModule", _vbProject.Object, 1);
            var qualifiedSelection = new Parsing.QualifiedSelection(selectedComponent, symbolSelection);

            // just for passing null reference exception            
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(1);
            context.SetupGet(c => c.Stop.Line).Returns(2);
            context.SetupGet(c => c.Stop.Text).Returns("Four");

            // simulate all the components and symbols   
            var member = new Parsing.QualifiedMemberName(selectedComponent, "fakeModule");
            const string identifierName = "Foo";
            AddDeclarationItem(context, symbolSelection, member, DeclarationType.Procedure, identifierName);
            AddDeclarationItem(context, new Parsing.Selection(1, 1, 1, 16), member, DeclarationType.Procedure);
            AddDeclarationItem(context, new Parsing.Selection(1, 1, 1, 1), member, DeclarationType.Module);

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
            var symbolSelection = new Parsing.Selection(4, 5, 4, 8);
            var selectedComponent = new Parsing.QualifiedModuleName("TestProject", "Module1", _vbProject.Object, 1);
            var qualifiedSelection = new Parsing.QualifiedSelection(selectedComponent, symbolSelection);

            // just for passing null reference exception            
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(-1);
            context.SetupGet(c => c.Stop.Line).Returns(-1);
            context.SetupGet(c => c.Stop.Text).Returns("Fake");

            // simulate all the components and symbols   
            IdentifierReference reference;
            var differentComponent = new Parsing.QualifiedModuleName("TestProject", "Module2", _vbProject.Object, 1);
            var differentMember = new Parsing.QualifiedMemberName(differentComponent, "Module2");
            AddDeclarationItem(context, new Parsing.Selection(4, 9, 4, 16), differentMember, DeclarationType.Variable,"FooTest");

            // add references to the Foo declaration item to simulate prod usage
            AddDeclarationItem(context, new Parsing.Selection(3, 5, 3, 8), differentMember, DeclarationType.Procedure, "Foo");
            var declarationItem = _listDeclarations[_listDeclarations.Count - 1];
            reference = new IdentifierReference(selectedComponent, "Foo", new Parsing.Selection(7, 5, 7, 11), false,context.Object, declarationItem);
            AddReference(declarationItem, reference);
            reference = new IdentifierReference(selectedComponent, "Foo", symbolSelection, false, context.Object,declarationItem);
            AddReference(declarationItem, reference);

            AddDeclarationItem(context, new Parsing.Selection(1, 1, 1, 1), differentMember, DeclarationType.Module, "Module2");
            var member = new Parsing.QualifiedMemberName(selectedComponent, "fakeModule");
            AddDeclarationItem(context, new Parsing.Selection(7, 5, 7, 11), member, DeclarationType.Procedure, "RunFoo");
            AddDeclarationItem(context, new Parsing.Selection(3, 5, 3, 9), member, DeclarationType.Procedure, "Main");
            AddDeclarationItem(context, new Parsing.Selection(1, 1, 1, 1), member, DeclarationType.Module, "Module1");

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

        [TestMethod]
        public void OnOkButtonClicked_GivinValidSettings_CodeModuleRenamed()
        {
            // arange
            const string moduleName = "Module1"; // inside this module the codeModule will be renamed
            const string projectName = "MainProject"; // this project will be treated as the one where we rename
            const string newName = "RubberDuckNewName";
            
            Mock<CodeModule> codeModule = new Mock<CodeModule>();
            codeModule.SetupProperty(x => x.Name);
            
            IList<VBComponent> componentList = new List<VBComponent>();
            componentList.Add(AddComponent("Class1", codeModule.Object).Object);
            componentList.Add(AddComponent(moduleName, codeModule.Object).Object);
            componentList.Add(AddComponent("Form1", codeModule.Object).Object);
            VBComponents vbComponnets = new CommonObjects.VbComponentsFake(componentList);

            IList<VBProject> projectList = new List<VBProject>();
            projectList.Add(AddProject("SecondProject", vbComponnets).Object);
            projectList.Add(AddProject(projectName, vbComponnets).Object);
            projectList.Add(AddProject("ThirdProject", vbComponnets).Object);
            VBProjects vbProjects = new CommonObjects.VbProjecstFake(projectList);

            _vbe.Setup(x => x.VBProjects).Returns(vbProjects);

            var symbolSelection = new Parsing.Selection(1, 1, 2, 2);
            var selectedComponent = new Parsing.QualifiedModuleName(projectName, moduleName, projectList.First(x => x.Name == projectName), 1);
            var qualifiedSelection = new Parsing.QualifiedSelection(selectedComponent, symbolSelection);
            var member = new Parsing.QualifiedMemberName(selectedComponent, moduleName);

            // just for passing null reference exception
            var context = new Mock<ParserRuleContext>();
            context.SetupGet(c => c.Start.Line).Returns(1);
            context.SetupGet(c => c.Stop.Line).Returns(2);
            context.SetupGet(c => c.Stop.Text).Returns("Four");

            // setting a declaration item as a procedure that will be renamed
            const string identifierName = moduleName;
            AddDeclarationItem(context, symbolSelection, member, DeclarationType.Module, identifierName);

            //act
            var presenter = new RenamePresenter(_vbe.Object, _view.Object, _declarations, qualifiedSelection);
            _view.Setup(x => x.Target).Returns(_declarations.Items.ToList().Find(x => x.ProjectName == projectName));
            _view.SetupGet(x => x.NewName).Returns(newName);
            _view.Raise(e => e.OkButtonClicked += null, EventArgs.Empty);
            
            //assert
            var project = _view.Object.Target.QualifiedName.QualifiedModuleName.Project;
            var actuallCodeModuleName = project.VBComponents.Cast<VBComponent>().First(x => x.Name == moduleName).CodeModule.Name;
            Assert.AreEqual(actuallCodeModuleName, newName);
        }
    }

}