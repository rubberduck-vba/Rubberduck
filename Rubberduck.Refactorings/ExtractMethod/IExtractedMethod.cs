using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractedMethod
    {
        Accessibility Accessibility { get; set; }
        string NewMethodCall();
        string MethodName { get; }
        string NewMethodNameBase { get; set; }
        IEnumerable<ExtractMethodParameter> Parameters { get; set; }
        ExtractMethodParameter ReturnValue { get; set; }
        bool SetReturnValue { get; set; }
    }
}
