namespace Rubberduck.Parsing.VBA
{
    //note: ordering of the members is important
    public enum ParserState
    {
        /// <summary>
        /// Parse has not been requested or has not started yet.
        /// </summary>
        Pending,
        /// <summary>
        /// Parse has started and is in the coordination phase.
        /// </summary>
        Started,
        /// <summary>
        /// Project references are being loaded into parser state.
        /// </summary>
        LoadingReference,
        /// <summary>
        /// Code from modified modules is being parsed.
        /// </summary>
        Parsing,
        /// <summary>
        /// Parse tree is waiting to be walked for identifier resolution.
        /// </summary>
        Parsed,
        /// <summary>
        /// Resolving declarations.
        /// </summary>
        ResolvingDeclarations,
        /// <summary>
        /// Resolved declarations.
        /// </summary>
        ResolvedDeclarations,
        /// <summary>
        /// Resolving identifier references.
        /// </summary>
        ResolvingReferences,
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Ready,
        /// <summary>
        /// The parser cannot run during that time (e.g. unit tests are running); any parse requests will be queued.
        /// </summary>
        Busy,
        /// <summary>
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error,
        /// <summary>
        /// Parsing completed, but identifier references could not be resolved for one or more modules.
        /// </summary>
        ResolverError,
        /// <summary>
        /// Unexpected exception has been encountered during a parse.
        /// </summary>
        UnexpectedError,
        /// <summary>
        /// This component doesn't need a state.  Use for built-in declarations.
        /// </summary>
        None,
    }
}
