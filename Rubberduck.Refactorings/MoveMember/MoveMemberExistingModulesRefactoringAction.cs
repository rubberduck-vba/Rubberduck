using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings
{
    public class MoveMemberExistingModulesRefactoringAction :  CodeOnlyRefactoringActionBase<MoveMemberModel>
    {
        private readonly IRewritingManager _rewritingManager;
        private readonly IMovedContentProviderFactory _contentProviderFactory;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberExistingModulesRefactoringAction(
                        IRewritingManager rewritingManager,
                        IMovedContentProviderFactory contentProviderFactory)
                : base(rewritingManager)
        {
            _rewritingManager = rewritingManager;
            _contentProviderFactory = contentProviderFactory;
            ContentProvider = _contentProviderFactory.CreateDefaultProvider();
        }

        public IMovedContentProvider ContentProvider { set; get; }

        public override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            GetSingleUseContentProvider(out var singleUseContentProvider);

            if (!model.TryFindApplicableStrategy(out var strategy)
                || !strategy.IsExecutableModel(model, out _))
            {
                return;
            }

            MoveMembers(model, strategy, rewriteSession, singleUseContentProvider);
        }

        public string NewModuleContent(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            GetSingleUseContentProvider(out var singleUseContentProvider);

            if (!model.TryFindApplicableStrategy(out var strategy))
            {
                return Resources.RubberduckUI.MoveMember_ApplicableStrategyNotFound;
            }


            strategy.RefactorRewrite(model, rewriteSession, _rewritingManager, singleUseContentProvider, out var newContent);

            return $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}{Environment.NewLine}{newContent}";
        }

        private void MoveMembers(MoveMemberModel model, IMoveMemberRefactoringStrategy strategy, IRewriteSession rewriteSession, IMovedContentProvider contentProvider)
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

        //Property ContentProvider may have had a 'preview' version Property Injected.  Always
        //restore the default ContentProvider after using the currently set reference.
        private void GetSingleUseContentProvider(out IMovedContentProvider singleUseContentProvider)
        {
            singleUseContentProvider = ContentProvider;
            ContentProvider = _contentProviderFactory.CreateDefaultProvider();
        }
    }
}
