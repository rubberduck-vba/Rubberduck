using System;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class WrapperMethodException : Exception
    {
        public WrapperMethodException(Exception inner)
            : base("COM wrapper method call threw an exception.", inner) { }
    }
}