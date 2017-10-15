using System;

namespace Rubberduck.Root
{
    public class InterceptedException : Exception
    {
        public InterceptedException(Exception inner)
            : this(inner.Message, inner) { }

        public InterceptedException(string message, Exception inner)
            : base(message, inner) { }
    }
}