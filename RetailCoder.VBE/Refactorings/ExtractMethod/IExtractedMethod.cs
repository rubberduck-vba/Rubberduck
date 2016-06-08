using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractedMethod
    {
        Accessibility Accessibility { get; set; }
        string NewMethodCall();
        string MethodName { get; set; }
        IEnumerable<ExtractedParameter> Parameters { get; set; }
        ExtractedParameter ReturnValue { get; set; }
        bool SetReturnValue { get; set; }
    }
}
