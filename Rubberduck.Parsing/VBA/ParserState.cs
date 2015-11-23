namespace Rubberduck.Parsing.VBA
{
    public enum ParserState
    {
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Ready,
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
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error
    }
}