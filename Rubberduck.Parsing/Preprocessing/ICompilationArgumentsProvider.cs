using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface ICompilationArgumentsProvider
    {
        VBAPredefinedCompilationConstants PredefinedCompilationConstants { get; }
        Dictionary<string, short> UserDefinedCompilationArguments(string projectId);
    }
}
