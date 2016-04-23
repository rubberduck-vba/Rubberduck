namespace Rubberduck.Parsing.VBA
{
    //note: ordering of the members is important
    public enum ParserState
    {
        /// <summary>
        /// Parse was requested but hasn't started yet.
        /// </summary>
        Pending,
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
        /// Resolving identifier references.
        /// </summary>
        Resolving,
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Ready,
        /// <summary>
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error,
        /// <summary>
        /// Parsing completed, but identifier references could not be resolved for one or more modules.
        /// </summary>
        ResolverError,
    }
}