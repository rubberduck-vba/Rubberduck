using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractedMethod : IExtractedMethod
    {
        public string MethodName { get; set; }
        public Accessibility Accessibility { get; set; }
        public bool SetReturnValue { get; set; }
        public ExtractedParameter ReturnValue { get; set; }
        public IEnumerable<ExtractedParameter> Parameters { get; set; }
        public string AsString()
        {
            string result = "" + MethodName;
            string argList;
            if (Parameters.Any())
            {
                argList = string.Join(", ", Parameters.Select(p => p.Name));
                result += " " + argList;
            }
            return result;
        }

    }
}
