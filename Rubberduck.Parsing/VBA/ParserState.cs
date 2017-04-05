using System;

namespace Rubberduck.Parsing.VBA
{
    //note: ordering of the members is important
    [Flags]
    public enum ParserState
    {
        /// <summary>
        /// Parse was requested but hasn't started yet.
        /// </summary>
        Pending = 1,
        /// <summary>
        /// Project references are being loaded into parser state.
        /// </summary>
        LoadingReference = 2,
        /// <summary>
        /// Code from modified modules is being parsed.
        /// </summary>
        Parsing = 4,
        /// <summary>
        /// Parse tree is waiting to be walked for identifier resolution.
        /// </summary>
        Parsed = 8,
        /// <summary>
        /// Resolving declarations.
        /// </summary>
        ResolvingDeclarations = 16,
        /// <summary>
        /// Resolved declarations.
        /// </summary>
        ResolvedDeclarations = 32,
        /// <summary>
        /// Resolving identifier references.
        /// </summary>
        ResolvingReferences = 64,
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Ready = 128,
        /// <summary>
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error = 256,
        /// <summary>
        /// Parsing completed, but identifier references could not be resolved for one or more modules.
        /// </summary>
        ResolverError = 512,
        /// <summary>
        /// This component doesn't need a state.  Use for built-in declarations.
        /// </summary>
        None = 1024,
    }
}
