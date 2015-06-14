using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;

namespace RubberduckTests.UI.Refactorings.ReorderParameters
{
    [TestClass]
    public class ReorderParametersRefactoringTests
    {
        private static Mock<VBE> _vbe;
        private static Mock<VBProject> _vbProject;
        private static Declarations _declarations;
        private static QualifiedModuleName _module;
        private static Mock<IReorderParametersView> _view;
        private List<Declaration> _listDeclarations;

        [TestInitialize]
        public void TestInitialization()
        {
            _vbe = new Mock<VBE>();
            _vbProject = new Mock<VBProject>();
            _declarations = new Declarations();
            _module = new QualifiedModuleName();
            _view = new Mock<IReorderParametersView>();
        }

        /// <summary>Common method for adding declaration with some default values</summary>
        private void AddDeclarationItem(IMock<ParserRuleContext> context,
            Selection selection,
            QualifiedMemberName? qualifiedName = null,
            DeclarationType declarationType = DeclarationType.Project,
            string identifierName = "identifierName")
        {
            Declaration declarationItem = null;
            var qualName = qualifiedName ?? new QualifiedMemberName(_module, "fakeModule");

            declarationItem = new Declaration(
                qualifiedName: qualName,
                parentScope: "module.proc",
                asTypeName: "asTypeName",
                isSelfAssigned: false,
                isWithEvents: false,
                accessibility: Accessibility.Public,
                declarationType: declarationType,
                context: context.Object,
                selection: selection
                );

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
    }
}