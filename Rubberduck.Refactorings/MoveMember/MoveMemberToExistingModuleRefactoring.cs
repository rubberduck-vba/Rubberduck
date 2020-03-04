using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToExistingModuleRefactoring : CodeOnlyRefactoringActionBase<MoveMemberModel>
    {
        private readonly IParseManager _parseManager;
        private readonly IRewritingManager _rewritingManager;
        private readonly IAddComponentService _addComponentService;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberToExistingModuleRefactoring(
                        IParseManager parseManager,
                        IRewritingManager rewritingManager,
                        IAddComponentService addComponentService)
                : base(rewritingManager)
        {
            _addComponentService = addComponentService;
            _parseManager = parseManager;
            _rewritingManager = rewritingManager;
        }

        public override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy) || !strategy.IsExecutableModel(model, out _))
            {
                return;
            }

            MoveMembers(model, strategy, rewriteSession);
        }

        private void MoveMembers(MoveMemberModel model, IMoveMemberRefactoringStrategy strategy, IRewriteSession rewriteSession)
        {
            try
            {
                strategy.RefactorRewrite(model, rewriteSession, _rewritingManager);
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
