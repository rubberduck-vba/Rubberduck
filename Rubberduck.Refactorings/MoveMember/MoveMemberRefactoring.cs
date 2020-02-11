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
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringTestAccess
    {
        MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction);
        string PreviewDestination(MoveMemberModel model);
    }

    public class MoveMemberRefactoring : InteractiveRefactoringBase<IMoveMemberPresenter, MoveMemberModel>, IMoveMemberRefactoringTestAccess
    {
        private readonly IMessageBox _messageBox;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseManager _parseManager;
        private readonly IRewritingManager _rewritingManager;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly ISelectionService _selectionService;

        private MoveMemberModel Model { set; get; } = null;

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
        }

        public MoveMemberModel TestUserInteractionOnly(Declaration target, Func<MoveMemberModel, MoveMemberModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        public string PreviewDestination(MoveMemberModel model)
        {
            var contentToMove = new ContentToMove();
            (bool strategyFound, IMoveMemberRewriteSession moveMemberSession) = RefactorRewrite(model, model.MoveRewritingManager.CheckOutCodePaneSession(), contentToMove);

            if (strategyFound)
            {
                if (model.Destination.IsExistingModule(out var destinationModule))
                {
                    var rewriter = moveMemberSession.CheckOutModuleRewriter(destinationModule.QualifiedModuleName);
                    return rewriter.GetText();
                }
                return contentToMove.AsSingleBlock;
            }

            return MoveMemberResources.ApplicableStrategyNotFound;
        }

        protected override MoveMemberModel InitializeModel(Declaration target)
        {
            if (target == null) { throw new TargetDeclarationIsNullException(); }

            Model = new MoveMemberModel(_declarationFinderProvider, RewritingManager, PreviewDestination);
            Model.DefineMove(target);
            return Model;
        }

        protected override void RefactorImpl(MoveMemberModel model)
        {
            if (!model.HasValidDestination)
            {
                _messageBox?.Message(MoveMemberResources.InvalidMoveDefinition);
                return;
            }

            Model = model;

            if (Model.Destination.IsExistingModule(out _))
            {
                SafeMoveMembers();
                return;
            }

            var suspendResult = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, SafeMoveMembers);
            var suspendOutcome = suspendResult.Outcome;
            if (suspendOutcome != SuspensionOutcome.Completed)
            {
                _logger.Warn($"AddModule: {Model.Destination.ModuleName} failed.");
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

        private static (bool strategyFound, IMoveMemberRewriteSession rewriteSession) RefactorRewrite(MoveMemberModel model, IExecutableRewriteSession session, IProvideNewContent contentToMove)
        {
            var strategy = model.Strategy;
            if (strategy is null)
            {
                return (false, new MoveMemberRewriteSession(session));
            }

            return (true, strategy.ModifyContent(model, session, contentToMove));
        }

        private void SafeMoveMembers()
        {
            ICodeModule newlyCreatedCodeModule = null;
            var newModulePostMoveSelection = new Selection();
            try
            {
                var contentToMove = new ContentToMove();
                (bool strategyFound, IMoveMemberRewriteSession moveMemberRewriteSession) = RefactorRewrite(Model, _rewritingManager.CheckOutCodePaneSession(), contentToMove);

                if (!strategyFound) { return; }

                if (!moveMemberRewriteSession.TryRewrite())
                {
                    PresentMoveMemberErrorMessage(BuildDefaultErrorMessage(Model.SelectedDeclarations.FirstOrDefault()));
                    return;
                }

                if (!Model.Destination.IsExistingModule(out _))
                {
                    //CreateNewModuleWithContent returns an ICodeModule reference to support setting the post-move Selection.
                    //Unable to use the ISelectionService after creating a module, since the
                    //new Component is not available via VBComponents until after a reparse
                    newlyCreatedCodeModule = CreateNewModule(contentToMove.AsSingleBlock, Model);
                    newModulePostMoveSelection = new Selection(newlyCreatedCodeModule.CountOfLines - contentToMove.CountOfLines + 1, 1);
                }
            }
            //TODO: Review these catches
            catch (MoveMemberUnsupportedMoveException unsupportedMove)
            {
                PresentMoveMemberErrorMessage(unsupportedMove.Message);
            }
            catch (RuntimeBinderException rbEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.SelectedDeclarations.FirstOrDefault())}: {rbEx.Message}");
            }
            catch (COMException comEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.SelectedDeclarations.FirstOrDefault())}: {comEx.Message}");
            }
            catch (ArgumentException argEx)
            {
                //This exception is often thrown when there is a rewrite conflict (e.g., try to insert where something's been deleted)
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.SelectedDeclarations.FirstOrDefault())}: {argEx.Message}");
            }
            catch (Exception unhandledEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.SelectedDeclarations.FirstOrDefault())}: {unhandledEx.Message}");
            }
            finally
            {
                if (newlyCreatedCodeModule != null)
                {
                    using (newlyCreatedCodeModule)
                    {
                        SetPostMoveSelection(newModulePostMoveSelection, newlyCreatedCodeModule);
                    }
                }
                else
                {
                    SetPostMoveSelection();
                }
            }
        }

        private static ICodeModule CreateNewModule(string newModuleContent, MoveMemberModel model)
        {
            ICodeModule codeModule = null;
            var vbProject = model.Source.Module.Project;
            using (var components = vbProject.VBComponents)
            {
                using (var newComponent = components.Add(model.Destination.ComponentType))
                {
                    newComponent.Name = model.Destination.ModuleName;
                    using (var newModule = newComponent.CodeModule)
                    {
                        //If VBE Option 'Require Variable Declaration' is set, then
                        //Option Explicit is included with a newly inserted Module...hence, the check
                        if (newModule.Content().Contains(MoveMemberResources.OptionExplicit))
                        {
                            newModule.InsertLines(newModule.CountOfLines, newModuleContent);
                        }
                        else
                        {
                            newModule.InsertLines(1, $"{MoveMemberResources.OptionExplicit}{Environment.NewLine}{Environment.NewLine}{newModuleContent}");
                        }
                        codeModule = newModule;
                    }
                }
            }
            return codeModule;
        }

        private void SetPostMoveSelection(Selection postMoveSelection = new Selection(), ICodeModule newlyCreatedCodeModule = null)
        {
            //The move/rewrite is done at this point, so do not bubble up any exceptions. 
            //If the user sees an exception, he may think that the the move failed
            try
            {
                if (newlyCreatedCodeModule != null)
                {
                    using (var codePane = newlyCreatedCodeModule.CodePane)
                    {
                        if (!codePane.IsWrappingNullReference)
                        {
                            codePane.Selection = postMoveSelection;
                        }
                    }
                    return;
                }
                if (Model.Destination.IsExistingModule(out var module))
                {
                    var destinationMembers = _declarationFinderProvider.DeclarationFinder.Members(module.QualifiedModuleName)
                        .Where(d => d.IsMember());

                    var lastPreMoveDestinationMember = destinationMembers.Where(d => d.IsMember()).OrderBy(d => d.Selection).LastOrDefault();

                    _selectionService.TrySetSelection(module.QualifiedModuleName, new Selection(lastPreMoveDestinationMember?.Context.Stop.Line ?? 1, 1));
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"{ex.Message}: {nameof(SetPostMoveSelection)} threw and exception");
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
            : base(String.Format(MoveMemberResources.UnsupportedMoveExceptionFormat, ToLocalizedString(declaration.DeclarationType) , declaration.IdentifierName))
        { }

        private static string ToLocalizedString(DeclarationType type)
            => RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);
    }
}
