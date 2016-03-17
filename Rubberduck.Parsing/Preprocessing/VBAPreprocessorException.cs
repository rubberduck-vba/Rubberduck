using System;

namespace Rubberduck.Parsing.Preprocessing
{
    [Serializable]
    public class VBAPreprocessorException : Exception
    {
        public VBAPreprocessorException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
