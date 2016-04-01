namespace Rubberduck.Parsing.VBA
{
    public enum ParserState
    {
        /// <summary>
        /// Parse was requested but hasn't started yet.
        /// </summary>
        Pending = 0,
        /// <summary>
        /// Project references are being loaded into parser state.
        /// </summary>
        LoadingReference = 1,
        /// <summary>
        /// Code from modified modules is being parsed.
        /// </summary>
        Parsing = 2,
        /// <summary>
        /// Parse tree is waiting to be walked for identifier resolution.
        /// </summary>
        Parsed = 3,
        /// <summary>
        /// Resolving identifier references.
        /// </summary>
        Resolving = 4,
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Ready = 5,
        /// <summary>
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error = 99,
        /// <summary>
        /// Parsing completed, but identifier references could not be resolved for one or more modules.
        /// </summary>
        ResolverError = 6,
    }
}