using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface IParseManager
    {
        event EventHandler<ParserStateEventArgs> StateChanged;
        event EventHandler<ParseProgressEventArgs> ModuleStateChanged;

        void OnParseRequested(object requestor);
        SuspensionResult OnSuspendParser(object requestor, IEnumerable<ParserState> allowedRunStates, Action busyAction, int millisecondsTimeout = -1);
        void MarkAsModified(QualifiedModuleName module);
    }
}