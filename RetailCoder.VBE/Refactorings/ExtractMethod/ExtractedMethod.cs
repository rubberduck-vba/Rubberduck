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
            string result;
            var argsList = string.Join(", ", Parameters.Select(p => p.Name));
            result = "" + MethodName + " " + argsList + " ";
            return result;
        }

    }
}
