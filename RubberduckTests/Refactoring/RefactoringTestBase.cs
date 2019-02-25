using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public abstract class RefactoringTestBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void NoActiveSelection_Throws()
        {
            var rewritingManager = new Mock<IRewritingManager>().Object;
            var refactoring = TestRefactoring(rewritingManager, null, initialSelection: null);

            Assert.Throws<NoActiveSelectionException>(() => refactoring.Refactor());
        }

        protected string RefactoredCode(string code, Selection selection, Type expectedException = null, bool setActiveSelection = false)
        {
            var vbe = TestVbe(code, out _);
            var componentName = vbe.SelectedVBComponent.Name;
            var refactored = RefactoredCode(vbe, componentName, selection, expectedException, setActiveSelection);
            return refactored[componentName];
        }

        protected IDictionary<string, string> RefactoredCode(string selectedComponentName, Selection selection, Type expectedException = null, bool setActiveSelection = false, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = TestVbe(modules);
            return RefactoredCode(vbe, selectedComponentName, selection, expectedException, setActiveSelection);
        }

        protected IDictionary<string, string> RefactoredCode(IVBE vbe, string selectedComponentName, Selection selection, Type expectedException = null, bool setActiveSelection = false)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var module = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .Single(declaration => declaration.IdentifierName == selectedComponentName)
                    .QualifiedModuleName;
                var qualifiedSelection = new QualifiedSelection(module, selection);

                var refactoring = setActiveSelection
                    ? TestRefactoring(rewritingManager, state, qualifiedSelection)
                    : TestRefactoring(rewritingManager, state);

                if (expectedException != null)
                {
                    Assert.Throws(expectedException, () => refactoring.Refactor(qualifiedSelection));
                }
                else
                {
                    refactoring.Refactor(qualifiedSelection);
                }

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        protected IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            QualifiedSelection? initialSelection = null
        )
        {
            var selectionService = MockedSelectionService(initialSelection);
            return TestRefactoring(rewritingManager, state, selectionService);
        }

        protected abstract IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            ISelectionService selectionService
        );

        protected virtual ISelectionService MockedSelectionService(QualifiedSelection? initialSelection)
        {
            QualifiedSelection? activeSelection = null;
            var selectionServiceMock = new Mock<ISelectionService>();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            selectionServiceMock.Setup(m => m.TrySetActiveSelection(It.IsAny<QualifiedSelection>()))
                .Returns(() => true).Callback((QualifiedSelection selection) => activeSelection = selection);
            return selectionServiceMock.Object;
        }

        protected virtual IVBE TestVbe(string content, out IVBComponent component, Selection? selection = null)
        {
            return MockVbeBuilder.BuildFromSingleStandardModule(content, out component, selection ?? default).Object;
        }

        protected virtual IVBE TestVbe(params (string componentName, string content, ComponentType componentType)[] modules)
        {
            return MockVbeBuilder.BuildFromModules(modules).Object;
        }
    }
}