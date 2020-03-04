using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberRefactoringAction : RefactoringActionWithSuspension<MoveMemberModel> //IRefactoringAction<MoveMemberModel>
    {
        private readonly IParseManager _parseManager;
        private readonly IRewritingManager _rewritingManager;
        private readonly IAddComponentService _addComponentService;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberRefactoringAction(
            IParseManager parseManager,
            IRewritingManager rewritingManager,
            IAddComponentService addComponentService)
                :base(parseManager, rewritingManager)
        {
            _addComponentService = addComponentService;
            _parseManager = parseManager;
            _rewritingManager = rewritingManager;
        }

        protected override bool RequiresSuspension(MoveMemberModel model)
                                    => !model.Destination.IsExistingModule(out _);

        protected override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            try
            {
                MoveMembers(model);
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

        private void MoveMembers(MoveMemberModel model)
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
                _addComponentService.AddComponentWithAttributes(
                    model.Source.Module.ProjectId, 
                    model.Destination.ComponentType, 
                    contentToMove.AsSingleBlock, 
                    componentName: model.Destination.ModuleName);
            }
        }
    }
}
