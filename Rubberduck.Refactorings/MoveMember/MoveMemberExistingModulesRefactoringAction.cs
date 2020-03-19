using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Refactorings.Rename;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberExistingModulesRefactoringAction : CodeOnlyRefactoringActionBase<MoveMemberModel>
    {
        private readonly IParseManager _parseManager;
        private readonly IRewritingManager _rewritingManager;
        private readonly RenameCodeDefinedIdentifierRefactoringAction _renameAction;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberExistingModulesRefactoringAction(
                        RenameCodeDefinedIdentifierRefactoringAction renameAction,
                        IParseManager parseManager,
                        IRewritingManager rewritingManager)
                : base(rewritingManager)
        {
            _renameAction = renameAction;
            _parseManager = parseManager;
            _rewritingManager = rewritingManager;
        }

        public override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy) || !strategy.IsExecutableModel(model, out _))
            {
                return;
            }

            model.RenameService = RenameService;

            MoveMembers(model, strategy, rewriteSession);
        }

        private void RenameService(Declaration declaration, string newName, IRewriteSession rewriteSession)
        {
            if (declaration.IdentifierName.IsEquivalentVBAIdentifierTo(newName)) { return; }

            var renameModel = new RenameModel(declaration)
            {
                NewName = newName,
            };

            _renameAction.Refactor(renameModel, rewriteSession);
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
