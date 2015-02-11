using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SourceControl
{
    public class SourceControlException : Exception
    {
        public SourceControlException() { }
        public SourceControlException(string message) : base(message) { }
        public SourceControlException(string message, Exception inner) : base(message, inner) { }
        protected SourceControlException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }

}
