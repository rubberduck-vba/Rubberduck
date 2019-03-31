using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface IParseManager
    {
        event EventHandler<ParserStateEventArgs> StateChanged;
        event EventHandler<ParseProgressEventArgs> ModuleStateChanged;

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
}