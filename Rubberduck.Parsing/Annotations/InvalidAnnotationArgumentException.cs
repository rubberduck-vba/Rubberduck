using System;

namespace Rubberduck.Parsing.Annotations
{
    [Serializable]
    public class InvalidAnnotationArgumentException : Exception
    {
        public InvalidAnnotationArgumentException(string message)
            : base(message)
        {
        }
    }
}
