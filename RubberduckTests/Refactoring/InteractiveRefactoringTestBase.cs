using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public abstract class InteractiveRefactoringTestBase<TPresenter, TModel> : RefactoringTestBase
        where TPresenter : class, IRefactoringPresenter<TModel>
        where TModel : class, IRefactoringModel
    {
        protected string RefactoredCode(string code, Selection selection, Func<TModel, TModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false)
        {
            var vbe = TestVbe(code, out _);
            var componentName = vbe.SelectedVBComponent.Name;
            var refactored = RefactoredCode(vbe, componentName, selection, presenterAdjustment, expectedException, executeViaActiveSelection);
            return refactored[componentName];
        }

        protected IDictionary<string, string> RefactoredCode(string selectedComponentName, Selection selection, Func<TModel, TModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = TestVbe(modules);
            return RefactoredCode(vbe, selectedComponentName, selection, presenterAdjustment, expectedException, executeViaActiveSelection);
        }

        protected IDictionary<string, string> RefactoredCode(IVBE vbe, string selectedComponentName, Selection selection, Func<TModel, TModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var module = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .Single(declaration => declaration.IdentifierName == selectedComponentName)
                    .QualifiedModuleName;
                var qualifiedSelection = new QualifiedSelection(module, selection);

                var refactoring = executeViaActiveSelection
                    ? TestRefactoring(rewritingManager, state, presenterAdjustment, qualifiedSelection)
                    : TestRefactoring(rewritingManager, state, presenterAdjustment);

                if (executeViaActiveSelection)
                {

                    if (expectedException != null)
                    {
                        Assert.Throws(expectedException, () => refactoring.Refactor());
                    }
                    else
                    {
                        refactoring.Refactor();
                    }
                }
                else
                {

                    if (expectedException != null)
                    {
                        Assert.Throws(expectedException, () => refactoring.Refactor(qualifiedSelection));
                    }
                    else
                    {
                        refactoring.Refactor(qualifiedSelection);
                    }
                }

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        protected string RefactoredCode(string code, string declarationName, DeclarationType declarationType, Func<TModel, TModel> presenterAdjustment, Type expectedException = null)
        {
            var vbe = TestVbe(code, out _);
            var componentName = vbe.SelectedVBComponent.Name;
            var refactored = RefactoredCode(vbe, declarationName, declarationType, presenterAdjustment, expectedException);
            return refactored[componentName];
        }

        protected IDictionary<string, string> RefactoredCode(string declarationName, DeclarationType declarationType, Func<TModel, TModel> presenterAdjustment, Type expectedException = null, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = TestVbe(modules);
            return RefactoredCode(vbe, declarationName, declarationType, presenterAdjustment, expectedException);
        }

        protected IDictionary<string, string> RefactoredCode(IVBE vbe, string declarationName, DeclarationType declarationType, Func<TModel, TModel> presenterAdjustment, Type expectedException = null)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Single(declaration => declaration.IdentifierName == declarationName);

                var refactoring = TestRefactoring(rewritingManager, state, presenterAdjustment);

                if (expectedException != null)
                {
                    Assert.Throws(expectedException, () => refactoring.Refactor(target));
                }
                else
                {
                    refactoring.Refactor(target);
                }

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            ISelectionService selectionService)
        {
            Func<TModel, TModel> presenterAdjustment = (model => model);
            return TestRefactoring(rewritingManager, state, presenterAdjustment, selectionService);
        }

        protected IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            Func<TModel, TModel> presenterAdjustment,
            QualifiedSelection? initialSelection = null)
        {
            var selectionService = MockedSelectionService(initialSelection);
            return TestRefactoring(rewritingManager, state, presenterAdjustment, selectionService);
        }

        protected IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            Func<TModel, TModel> presenterAdjustment,
            ISelectionService selectionService)
        {
            var factory = SetupFactory(presenterAdjustment);
            return TestRefactoring(rewritingManager, state, factory.Object, selectionService);
        }

        protected abstract IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            IRefactoringPresenterFactory factory,
            ISelectionService selectionService);

        protected virtual Mock<IRefactoringPresenterFactory> SetupFactory(Func<TModel, TModel> adjustedModel)
        {
            var presenter = new Mock<TPresenter>();

            var factory = new Mock<IRefactoringPresenterFactory>();
            factory.Setup(f => f.Create<TPresenter, TModel>(It.IsAny<TModel>()))
                .Callback((TModel model) => presenter.Setup(p => p.Show()).Returns(() => adjustedModel(model)))
                .Returns(presenter.Object);
            return factory;
        }
    }
}