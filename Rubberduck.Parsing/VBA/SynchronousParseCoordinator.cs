using System;
using System.Threading;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousParseCoordinator : ParseCoordinator
    {
        public SynchronousParseCoordinator(
            RubberduckParserState state,
            IParsingStageService parsingStageService,
            IParsingCacheService parsingCacheService,
            IProjectManager projectManager,
            IParserStateManager parserStateManager,
            IRewritingManager rewritingManager = null) 
        :base(
            state,
            parsingStageService,
            parsingCacheService,
            projectManager,
            parserStateManager,
            rewritingManager)
        {}

        public override void BeginParse(object sender)
        {
            ParseInternal(CurrentCancellationTokenSource.Token);
        }

        public void Parse(CancellationTokenSource token)
        {
            SetSavedCancellationTokenSource(token);
            ParseInternal(token.Token);
        }

        /// <summary>
        /// For the use of tests and reflection API only
        /// </summary>
        private void SetSavedCancellationTokenSource(CancellationTokenSource tokenSource)
        {
            var oldTokenSource = CurrentCancellationTokenSource;
            CurrentCancellationTokenSource = tokenSource;

            oldTokenSource?.Cancel();
            oldTokenSource?.Dispose();
        }

        protected void ParseInternal(CancellationToken token)
        {
            var lockTaken = false;
            try
            {
                if (!ParsingSuspendLock.IsWriteLockHeld)
                {
                    ParsingSuspendLock.EnterReadLock();
                }
                Monitor.Enter(ParsingRunSyncObject, ref lockTaken);
                ParseAllInternal(this, token);
            }
            catch (OperationCanceledException)
            {
                //This is the point to which the cancellation should break.
            }
            finally
            {
                if (lockTaken)
                {
                    Monitor.Exit(ParsingRunSyncObject);
                }
                if (ParsingSuspendLock.IsReadLockHeld)
                {
                    ParsingSuspendLock.ExitReadLock();
                }
            }
        }
    }
}
