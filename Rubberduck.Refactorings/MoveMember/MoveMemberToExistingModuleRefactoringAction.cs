using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings
{
    public class MoveMemberToExistingModuleRefactoringAction :  CodeOnlyRefactoringActionBase<MoveMemberModel>
    {
        private readonly IRewritingManager _rewritingManager;
        private readonly IMovedContentProviderFactory _contentProviderFactory;
        private readonly IMoveMemberStrategyFactory _strategyFactory;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberToExistingModuleRefactoringAction(
                        IRewritingManager rewritingManager,
                        IMovedContentProviderFactory contentProviderFactory,
                        IMoveMemberStrategyFactory strategyFactory)
                : base(rewritingManager)
        {
            _rewritingManager = rewritingManager;
            _contentProviderFactory = contentProviderFactory;
            _strategyFactory = strategyFactory;
        }

        public override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            if (!model.TryGetStrategy(out var strategy))
            {
                throw new MoveMemberUnsupportedMoveException(Resources.RubberduckUI.MoveMember_ApplicableStrategyNotFound);
            }

            if (!strategy.IsExecutableModel(model, out var msg))
            {
                throw new MoveMemberUnsupportedMoveException(msg);
            }

            MoveMembers(model, strategy, rewriteSession, _contentProviderFactory.CreateDefaultProvider());
        }

        private void MoveMembers(MoveMemberModel model, IMoveMemberRefactoringStrategy strategy, IRewriteSession rewriteSession, INewContentProvider contentProvider)
        {
            var newModuleContent = string.Empty;
            try
            {
                strategy.RefactorRewrite(model, rewriteSession, _rewritingManager, contentProvider, out newModuleContent);
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
        }
    }
}
