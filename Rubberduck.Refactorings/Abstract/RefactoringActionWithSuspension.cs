using System.Runtime.ExceptionServices;
using NLog;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.Refactorings
{
    public abstract class RefactoringActionWithSuspension<TModel> : RefactoringActionBase<TModel>
        where TModel : class, IRefactoringModel
    {
        private readonly IParseManager _parseManager;

        private readonly Logger _logger;

        protected RefactoringActionWithSuspension(IParseManager parseManager, IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _parseManager = parseManager;
            _logger = LogManager.GetLogger(GetType().FullName);
        }

        protected abstract bool RequiresSuspension(TModel model);

        public override void Refactor(TModel model)
        {
            if (RequiresSuspension(model))
            {
                RefactorWithSuspendedParser(model);
            }
            else
            {
                base.Refactor(model);
            }
        }

        private void RefactorWithSuspendedParser(TModel model)
        {
            var suspendResult = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, () => base.Refactor(model));
            var suspendOutcome = suspendResult.Outcome;
            if (suspendOutcome != SuspensionOutcome.Completed)
            {
                if ((suspendOutcome == SuspensionOutcome.UnexpectedError || suspendOutcome == SuspensionOutcome.Canceled)
                    && suspendResult.EncounteredException != null)
                {
                    ExceptionDispatchInfo.Capture(suspendResult.EncounteredException).Throw();
                    return;
                }

                _logger.Warn($"{GetType().Name} failed because a parser suspension request could not be fulfilled.  The request's result was '{suspendResult.ToString()}'.");
                throw new SuspendParserFailureException();
            }
        }
    }
}