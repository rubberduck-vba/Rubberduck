using Antlr4.Runtime;
using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
            IUiDispatcher uiDispatcher)
            : base(rewritingManager, selectionService, factory, uiDispatcher)
                  
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _messageBox = messageBox;
            _rewritingManager = rewritingManager;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _selectionService = selectionService;
            _moveMemberFactory = new MoveMemberObjectsFactory(declarationFinderProvider);
        }

        public MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        public string PreviewModuleContent(MoveMemberModel model, PreviewModule previewModule)
        {
            //If there are no declarations selected to move, preview the module's existing content
            //if (!model.SelectedDeclarations.Any())
            //{
            //    if (previewModule == PreviewModule.Source)
            //    {
            //        return PreviewExistingContent(model, model.Source.QualifiedModuleName);
            //    }

            //    if (previewModule == PreviewModule.Destination)
            //    {
            //        if (model.Destination.IsExistingModule(out var destination))
            //        {
            //            return PreviewExistingContent(model, destination.QualifiedModuleName);
            //        }

            //        return $"{Tokens.Option} {Tokens.Explicit}";
            //    }
            //}

            //if (!model.Strategy.IsApplicable(model))
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
            var preview =  rewriter.GetText(maxConsecutiveNewLines: 3);
            return preview;
        }

        private string PreviewExistingContent(MoveMemberModel model, QualifiedModuleName qualifiedModuleName)
        {
            var session = _rewritingManager.CheckOutCodePaneSession();
            var sourceRewriter = session.CheckOutModuleRewriter(qualifiedModuleName);
            return sourceRewriter.GetText();
        }

        protected override MoveMemberModel InitializeModel(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            var model = new MoveMemberModel(target, _declarationFinderProvider, PreviewModuleContent, _moveMemberFactory);
            return model;
        }

        protected override void RefactorImpl(MoveMemberModel model)
        {
            if (!model.HasValidDestination)
            {
                _messageBox?.Message(MoveMemberResources.InvalidMoveDefinition);
                return;
            }

            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out _))
            {
                _messageBox?.Message(MoveMemberResources.ApplicableStrategyNotFound);
                return;
            }

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

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selected = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selected.IsMember()
                || selected.IsModuleConstant()
                || selected.IsField())
            {
                return selected;
            }

            return null;
        }

        private void MoveMembers(MoveMemberModel model)
        {
            try
            {
                if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy))
                {
                    return;
                }

                var contentToMove = new ContentToMove();
                var moveMemberRewriteSession = strategy.RefactorRewrite(model, _rewritingManager, contentToMove);

                if (!moveMemberRewriteSession.TryRewrite())
                {
                    PresentMoveMemberErrorMessage(BuildDefaultErrorMessage(model.SelectedDeclarations.FirstOrDefault()));
                    return;
                }

                if (!model.Destination.IsExistingModule(out _))
                {
                    CreateNewModule(contentToMove.AsSingleBlock, model);
                }
            }
            //TODO: Review these catches
            catch (MoveMemberUnsupportedMoveException unsupportedMove)
            {
                PresentMoveMemberErrorMessage(unsupportedMove.Message);
            }
            catch (RuntimeBinderException rbEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(model.SelectedDeclarations.FirstOrDefault())}: {rbEx.Message}");
            }
            catch (COMException comEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(model.SelectedDeclarations.FirstOrDefault())}: {comEx.Message}");
            }
            catch (ArgumentException argEx)
            {
                //This exception is often thrown when there is a rewrite conflict (e.g., try to insert where something's been deleted)
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(model.SelectedDeclarations.FirstOrDefault())}: {argEx.Message}");
            }
            catch (Exception unhandledEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(model.SelectedDeclarations.FirstOrDefault())}: {unhandledEx.Message}");
            }
        }

        private void CreateNewModule(string newModuleContent, MoveMemberModel model)
        {
            var targetProject = model.Source.Module.Project;
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

        private void PresentMoveMemberErrorMessage(string errorMsg)
        {
            _messageBox?.NotifyWarn(errorMsg, MoveMemberResources.Caption);
        }

        private string BuildDefaultErrorMessage(Declaration target)
        {
            return string.Format(MoveMemberResources.DefaultErrorMessageFormat, target?.IdentifierName ?? MoveMemberResources.InvalidMoveDefinition);
        }
    }

    //TODO: Are there any tests checking for this exception?
    [Serializable]
    class MoveMemberUnsupportedMoveException : Exception
    {
        public MoveMemberUnsupportedMoveException() { }

        public MoveMemberUnsupportedMoveException(Declaration declaration)
            : base(String.Format(MoveMemberResources.UnsupportedMoveExceptionFormat, 
                        ToLocalizedString(declaration?.DeclarationType ?? DeclarationType.Member), 
                        declaration?.IdentifierName ?? ToLocalizedString(DeclarationType.Member)))
        { }

        private static string ToLocalizedString(DeclarationType type)
            => RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);
    }
}
