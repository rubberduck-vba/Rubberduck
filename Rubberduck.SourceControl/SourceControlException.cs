using System;
using System.Runtime.Serialization;

namespace Rubberduck.SourceControl
{
    [Serializable]
    public class SourceControlException : Exception
    {
        public SourceControlException() { }
        public SourceControlException(string message) : base(message) { }
        public SourceControlException(string message, Exception inner) : base(message, inner) { }

        protected SourceControlException(SerializationInfo info,StreamingContext context)
            : base(info, context) { }
    }

}
