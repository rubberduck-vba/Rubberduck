using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface IParseManager : IParserStatusProvider
    {
        /// <summary>
        /// Requests reparse.
        /// </summary>
        /// <param name="requestor">The object requesting a reparse.</param>
        void OnParseRequested(object requestor);

        /// <summary>
        /// Requests cancellation of the current parse.
        /// Always request a parse after requesting a cancellation.
        /// Failing to do so leaves RD in an undefined state.
        /// </summary>
        /// <param name="requestor">The object requesting a reparse.</param>
        void OnParseCancellationRequested(object requestor);

        /// <summary>
        /// Suspends the parser for the action provided after the current parse has finished.
        /// Any incoming parse requests will be executed afterwards.
        /// </summary>
        /// <param name="requestor">The object requesting a reparse.</param>
        /// <param name="allowedRunStates">The states in which the action may be performed.</param>
        /// <param name="busyAction">The action to perform.</param> 
        /// <param name="millisecondsTimeout">The timeout for waiting on the current parse to finish.</param> 
        SuspensionResult OnSuspendParser(object requestor, IEnumerable<ParserState> allowedRunStates, Action busyAction, int millisecondsTimeout = -1);
        void MarkAsModified(QualifiedModuleName module);
    }

    public enum SuspensionOutcome
    {
        /// <summary>
        /// The busy action has been queued but has not run yet.
        /// </summary>
        Pending,
        /// <summary>
        /// The busy action was completed successfully.
        /// </summary>
        Completed,
        /// <summary>
        /// The busy action could not executed because it timed out when
        /// attempting to obtain a suspension lock. The timeout is 
        /// governed by the MillisecondsTimeout argument.
        /// </summary>
        TimedOut,
        /// <summary>
        /// The parser arrived to one of states that wasn't listed in the 
        /// AllowedRunStates specified by the requestor (e.g. an error state)
        /// and thus the busy action was not executed.
        /// </summary>
        IncompatibleState,
        /// <summary>
        /// Indicates that the suspension request cannot be made because there 
        /// is no handler for it. This points to a bug in the code.
        /// </summary>
        NotEnabled,
        /// <summary>
        /// The suspend action has thrown an OperationCanceledException.
        /// </summary>
        Canceled,
        /// <summary>
        /// We already hold a read lock to the suspension lock; this indicates a bug in code.
        /// </summary>
        ReadLockAlreadyHeld,
        /// <summary>
        /// An unexpected error; usually indicates a bug in code.
        /// </summary>
        UnexpectedError
    }

    public readonly struct SuspensionResult
    {
        public SuspensionResult(SuspensionOutcome outcome, Exception encounteredException = null)
        {
            Outcome = outcome;
            EncounteredException = encounteredException;
        }

        public SuspensionOutcome Outcome { get; }
        public Exception EncounteredException { get; }
    }
}