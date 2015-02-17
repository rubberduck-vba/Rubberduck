using System;

namespace Rubberduck.VBA
{
    /// <summary>
    /// An exception thrown by an <c>IParseTreeListener</c> implementation 
    /// that does not need to traverse an entire parse tree.
    /// </summary>
    public class WalkerCancelledException : Exception
    {
        public WalkerCancelledException()
            : base("Tree walker was cancelled by listener.")
        { }
    }
}