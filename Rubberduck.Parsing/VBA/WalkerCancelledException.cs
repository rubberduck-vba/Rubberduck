using System;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// An exception thrown by an <c>IParseTreeListener</c> implementation 
    /// that does not need to traverse an entire parse tree.
    /// </summary>
    [Serializable]
    public class WalkerCancelledException : Exception
    {
        public WalkerCancelledException()
            : base("Tree walker was cancelled by listener.")
        { }
    }
}