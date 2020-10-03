using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
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

        protected IDictionary<string, string> RefactoredCode(IVBE vbe, string declarationName, DeclarationType declarationType, Func<TModel, TModel> presenterAdjustment, Type expectedException = null, bool extractAllProjects = false)
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

                if (extractAllProjects)
                {
                    return state.ProjectsProvider.Components()
                        .Select(tpl => tpl.Component)
                        .ToDictionary(component => component.Name, component => component.CodeModule.Content());
                }
                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }

        protected TModel InitialModel(string code, Selection selection, bool executeViaActiveSelection = false)
        {
            var vbe = TestVbe(code, out _);
            var componentName = vbe.SelectedVBComponent.Name;
            return InitialModel(vbe, componentName, selection, executeViaActiveSelection);
        }

        protected TModel InitialModel(string selectedComponentName, Selection selection, bool executeViaActiveSelection = false, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = TestVbe(modules);
            return InitialModel(vbe, selectedComponentName, selection, executeViaActiveSelection);
        }

        protected TModel InitialModel(IVBE vbe, string selectedComponentName, Selection selection, bool executeViaActiveSelection = false)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                TModel initialModel = null;
                Func<TModel, TModel> exfiltrationAction = model =>
                {
                    initialModel = model;
                    throw new RefactoringAbortedException();
                };

                var module = state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .Single(declaration => declaration.IdentifierName == selectedComponentName)
                    .QualifiedModuleName;
                var qualifiedSelection = new QualifiedSelection(module, selection);

                var refactoring = executeViaActiveSelection
                    ? TestRefactoring(rewritingManager, state, exfiltrationAction, qualifiedSelection)
                    : TestRefactoring(rewritingManager, state, exfiltrationAction);

                try
                {
                    if (executeViaActiveSelection)
                    {
                        refactoring.Refactor();
                    }
                    else
                    {
                        refactoring.Refactor(qualifiedSelection);
                    }
                }
                catch (RefactoringAbortedException)
                {}

                return initialModel;
            }
        }

        protected TModel InitialModel(string code, string declarationName, DeclarationType declarationType)
        {
            var vbe = TestVbe(code, out _);
            var componentName = vbe.SelectedVBComponent.Name;
            return InitialModel(vbe, declarationName, declarationType);
        }

        protected TModel InitialModel(string declarationName, DeclarationType declarationType, params (string componentName, string content, ComponentType componentType)[] modules)
        {
            var vbe = TestVbe(modules);
            return InitialModel(vbe, declarationName, declarationType);
        }

        protected TModel InitialModel(IVBE vbe, string declarationName, DeclarationType declarationType)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(declarationType)
                    .Single(declaration => declaration.IdentifierName == declarationName);

                TModel initialModel = null;
                Func<TModel,TModel> exfiltrationAction = model =>
                {
                    initialModel = model;
                    throw new RefactoringAbortedException();
                };

                var refactoring = TestRefactoring(rewritingManager, state, exfiltrationAction);

                try
                {
                    refactoring.Refactor(target);
                }
                catch (RefactoringAbortedException)
                {}

                return initialModel;
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
            var userInteraction = TestUserInteraction(presenterAdjustment);
            return TestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }

        protected IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            IRefactoringPresenterFactory factory,
            ISelectionService selectionService)
        {
            var userInteraction = TestUserInteraction(factory);
            return TestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }

        protected abstract IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<TPresenter, TModel> userInteraction,
            ISelectionService selectionService);

        protected RefactoringUserInteraction<TPresenter, TModel> TestUserInteraction(Func<TModel, TModel> presenterAdjustment)
        {
            var factory = TestPresenterFactory(presenterAdjustment);
            return TestUserInteraction(factory.Object);
        }

        protected RefactoringUserInteraction<TPresenter, TModel> TestUserInteraction(IRefactoringPresenterFactory factory)
        {
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());

            return new RefactoringUserInteraction<TPresenter, TModel>(factory, uiDispatcherMock.Object);
        }

        protected virtual Mock<IRefactoringPresenterFactory> TestPresenterFactory(Func<TModel, TModel> adjustedModel)
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