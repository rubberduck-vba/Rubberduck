using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface ICompilationArgumentsProvider
    {
        Dictionary<string, short> UserDefinedCompilationArguments(string projectId);
    }
}
