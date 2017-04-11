using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IParserStateManager
    {
        ParserState OverallParserState { get; }
        ParserState GetModuleState(QualifiedModuleName module);

        void SetModuleState(QualifiedModuleName module, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true);
        void SetModuleStates(IReadOnlyCollection<QualifiedModuleName> modules, ParserState parserState, CancellationToken token, bool evaluateOverallParserState = true);
        void EvaluateOverallParserState(CancellationToken token);
        void SetStatusAndFireStateChanged(object requestor, ParserState status, CancellationToken token);
    }
}
