using System.Runtime.InteropServices;
using Source = Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Registration;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.ParserStateGuid)
    ]
    public enum ParserState
    {
        Pending = Source.ParserState.Pending,
        Started = Source.ParserState.Started,
        LoadingReference = Source.ParserState.LoadingReference,
        Parsing = Source.ParserState.Parsing,
        Parsed = Source.ParserState.Parsed,
        ResolvingDeclarations = Source.ParserState.ResolvingDeclarations,
        ResolvedDeclarations = Source.ParserState.ResolvedDeclarations,
        ResolvingReferences = Source.ParserState.ResolvingReferences,
        Ready = Source.ParserState.Ready,
        Error = Source.ParserState.Error,
        ResolverError = Source.ParserState.ResolverError,
        UnexpectedError = Source.ParserState.UnexpectedError
    }
}
