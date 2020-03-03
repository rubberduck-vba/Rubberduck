using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Utility;
using System;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringTestAccess
    {
        MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction);
        string PreviewModuleContent(MoveMemberModel model, PreviewModule previewModule);
    }

    public class MoveMemberRefactoring : InteractiveRefactoringBase<IMoveMemberPresenter, MoveMemberModel>, IMoveMemberRefactoringTestAccess
    {
        private readonly IMessageBox _messageBox;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseManager _parseManager;
        private readonly IRewritingManager _rewritingManager;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly ISelectionService _selectionService;
        private readonly IProjectsProvider _projectsProvider;

        private MoveMemberObjectsFactory _moveMemberFactory;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberRefactoring(
            IDeclarationFinderProvider declarationFinderProvider,
            IParseManager parseManager,
            IMessageBox messageBox,
            IRefactoringPresenterFactory factory,
            IRewritingManager rewritingManager,
            ISelectionService selectionService,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IProjectsProvider projectsProvider,
            IUiDispatcher uiDispatcher)
            : base(rewritingManager, selectionService, factory, uiDispatcher)

        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _messageBox = messageBox;
            _rewritingManager = rewritingManager;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _selectionService = selectionService;
            _projectsProvider = projectsProvider;
            _moveMemberFactory = new MoveMemberObjectsFactory(declarationFinderProvider);
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selected = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selected.IsMember()
                || selected.IsModuleConstant()
                || (selected.IsField() && !selected.HasPrivateAccessibility()))
            {
                return selected;
            }

            return null;
        }

        public MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        public string PreviewModuleContent(MoveMemberModel model, PreviewModule previewModule)
        {
            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy))
            {
                return MoveMemberResources.ApplicableStrategyNotFound;
            }

            var isExistingDestination = model.Destination.IsExistingModule(out var destinationModule);
            if (previewModule == PreviewModule.Destination && !isExistingDestination)
            {
                var content = strategy.NewDestinationModuleContent(model, _rewritingManager, new ContentToMove()).AsSingleBlockWithinDemarcationComments();

                return $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}{Environment.NewLine}{content}";
            }

            var previewSession = strategy.RefactorRewrite(model, _rewritingManager, new ContentToMove(), true);

            var qmnToPreview = previewModule == PreviewModule.Destination
                ? destinationModule.QualifiedModuleName
                : model.Source.QualifiedModuleName;

            var rewriter = previewSession.CheckOutModuleRewriter(qmnToPreview);
            var preview = rewriter.GetText(maxConsecutiveNewLines: 3);
            return preview;
        }

        protected override MoveMemberModel InitializeModel(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            var model = new MoveMemberModel(target, _declarationFinderProvider, PreviewModuleContent, _moveMemberFactory);
            return model;
        }

        //https://github.com/rubberduck-vba/Rubberduck/pull/5387
        //TODO: Update once #5387 is merged to eliminate suspension code
        protected override void RefactorImpl(MoveMemberModel model)
        {
            if (model.Destination.IsExistingModule(out _))
            {
                MoveMembers(model);
                return;
            }

            var suspendResult = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, () => MoveMembers(model));
            var suspendOutcome = suspendResult.Outcome;
            if (suspendOutcome != SuspensionOutcome.Completed)
            {
                if ((suspendOutcome == SuspensionOutcome.UnexpectedError || suspendOutcome == SuspensionOutcome.Canceled)
                    && suspendResult.EncounteredException != null)
                {
                    ExceptionDispatchInfo.Capture(suspendResult.EncounteredException).Throw();
                    return;
                }

                _logger.Warn($"{nameof(MoveMembers)} failed because a parser suspension request could not be fulfilled.  The request's result was '{suspendResult.ToString()}'.");
                throw new SuspendParserFailureException();
            }
        }

        private void MoveMembers(MoveMemberModel model)
        {
            try
            {
                if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy) || !strategy.IsExecutableModel(model, out _))
                {
                    return;
                }

                var contentToMove = new ContentToMove();
                var moveMemberRewriteSession = strategy.RefactorRewrite(model, _rewritingManager, contentToMove);

                if (!moveMemberRewriteSession.TryRewrite())
                {
                    throw new RewriteFailedException(moveMemberRewriteSession);
                }

                if (!model.Destination.IsExistingModule(out _))
                {
                    CreateNewModule(contentToMove.AsSingleBlock, model);
                }
            }
            catch (MoveMemberUnsupportedMoveException unsupportedMove)
            {
                _logger.Warn($"{nameof(MoveMembers)} {nameof(MoveMemberUnsupportedMoveException)} {unsupportedMove.Message}");
            }
            catch (RuntimeBinderException rbEx)
            {
                _logger.Warn($"{nameof(MoveMembers)} {nameof(RuntimeBinderException)} {rbEx.Message}");
            }
            catch (COMException comEx)
            {
                _logger.Warn($"{nameof(MoveMembers)} {nameof(COMException)} {comEx.Message}");
            }
            catch (ArgumentException argEx)
            {
                //This exception is often thrown when there is a rewrite conflict (e.g., try to insert where something's been deleted)
                _logger.Warn($"{nameof(MoveMembers)} {nameof(ArgumentException)} {argEx.Message}");
            }
            catch (Exception unhandledEx)
            {
                _logger.Warn($"{nameof(MoveMembers)} {nameof(Exception)} {unhandledEx.Message}");
            }
        }

        //https://github.com/rubberduck-vba/Rubberduck/pull/5387
        //TODO: Update once #5387 is merged
        private void CreateNewModule(string newModuleContent, MoveMemberModel model)
        {
            var targetProject = _projectsProvider.Project(model.Source.Module.ProjectId);
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            using (var components = targetProject.VBComponents)
            {
                using (var newComponent = components.Add(model.Destination.ComponentType))
                {
                    newComponent.Name = model.Destination.ModuleName;
                    using (var newModule = newComponent.CodeModule)
                    {
                        //If VBE Option 'Require Variable Declaration' is set, then
                        //Option Explicit is included with a newly inserted Module...hence, the check
                        var optionExplicit = $"{Tokens.Option} {Tokens.Explicit}";
                        if (newModule.Content().Contains(optionExplicit))
                        {
                            newModule.InsertLines(newModule.CountOfLines, newModuleContent);
                            return;
                        }
                        newModule.InsertLines(1, $"{optionExplicit}{Environment.NewLine}{Environment.NewLine}{newModuleContent}");
                    }
                }
            }
        }
    }
}
