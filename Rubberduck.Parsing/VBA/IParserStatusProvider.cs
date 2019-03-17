using System;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public interface IParserStatusProvider
    {
        event EventHandler<ParserStateEventArgs> StateChanged;
        event EventHandler<ParseProgressEventArgs> ModuleStateChanged;

        ParserState Status { get; }
        ParserState GetModuleState(QualifiedModuleName module);
    }
}