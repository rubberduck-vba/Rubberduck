using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToNewModuleRefactoringAction : RefactoringActionWithSuspension<MoveMemberModel>
    {
        private readonly IRewritingManager _rewritingManager;
        private readonly IAddComponentService _addComponentService;
        private readonly IMovedContentProviderFactory _contentProviderFactory;
        private readonly IMoveMemberStrategyFactory _strategyFactory;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberToNewModuleRefactoringAction(
                        IParseManager parseManager,
                        IRewritingManager rewritingManager,
                        IMovedContentProviderFactory contentProviderFactory,
                        IMoveMemberStrategyFactory strategyFactory,
                        IAddComponentService addComponentService)
                : base(parseManager, rewritingManager)
        {
            _rewritingManager = rewritingManager;
            _addComponentService = addComponentService;
            _contentProviderFactory = contentProviderFactory;
            _strategyFactory = strategyFactory;
        }

        protected override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            var strategy = _strategyFactory.Create(model.MoveEndpoints);

            if (!strategy.IsExecutableModel(model, out var msg))
            {
                throw new MoveMemberUnsupportedMoveException(msg);
            }

            var newContent = MoveMembers(model, strategy, rewriteSession, _contentProviderFactory.CreateDefaultProvider());

            var optionExplicit = $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}";

            _addComponentService.AddComponentWithAttributes(
                                        model.Source.Module.ProjectId,
                                        model.Destination.ComponentType,
                                        $"{optionExplicit}{Environment.NewLine}{newContent}",
                                        componentName: model.Destination.ModuleName);
        }

        protected override bool RequiresSuspension(MoveMemberModel model) => true;

        private string MoveMembers(MoveMemberModel model, IMoveMemberRefactoringStrategy strategy, IRewriteSession rewriteSession, INewContentProvider contentProvider)
        {
            var newModuleContent = string.Empty;
            try
            {
                strategy.RefactorRewrite(model, rewriteSession, _rewritingManager, contentProvider, out newModuleContent);
                return newModuleContent;
            }
            catch (MoveMemberUnsupportedMoveException unsupportedMove)
            {
                _logger.Warn($"{nameof(MoveMembers)} {nameof(MoveMemberUnsupportedMoveException)} {unsupportedMove.Message}");
                throw;
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
            return newModuleContent;
        }
    }
}
