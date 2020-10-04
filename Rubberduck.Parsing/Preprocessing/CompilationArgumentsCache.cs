using System.Collections.Generic;
using Rubberduck.InternalApi.Extensions;

namespace Rubberduck.Parsing.PreProcessing
{
    public class CompilationArgumentsCache : ICompilationArgumentsCache
    {
        private readonly ICompilationArgumentsProvider _provider;
        private readonly Dictionary<string,Dictionary<string,short>> _compilationArguments = new Dictionary<string, Dictionary<string, short>>();
        private readonly HashSet<string> _projectsWhoseCompilationArgumentsChanged = new HashSet<string>();

        public CompilationArgumentsCache(ICompilationArgumentsProvider compilationArgumentsProvider)
        {
            _provider = compilationArgumentsProvider;
        }

        public VBAPredefinedCompilationConstants PredefinedCompilationConstants =>
            _provider.PredefinedCompilationConstants;

        public Dictionary<string, short> UserDefinedCompilationArguments(string projectId)
        {
            return _compilationArguments.TryGetValue(projectId, out var compilatioarguments)
                ? compilatioarguments
                : new Dictionary<string, short>();
        }

        public void ReloadCompilationArguments(IEnumerable<string> projectIds)
        {
            foreach (var projectId in projectIds)
            {
                var oldCompilationArguments = UserDefinedCompilationArguments(projectId);
                ReloadCompilationArguments(projectId);
                var newCompilationArguments = UserDefinedCompilationArguments(projectId);
                if (!newCompilationArguments.HasEqualContent(oldCompilationArguments))
                {
                    _projectsWhoseCompilationArgumentsChanged.Add(projectId);
                }
            }
        }

        private void ReloadCompilationArguments(string projectId)
        {
            _compilationArguments[projectId] = _provider.UserDefinedCompilationArguments(projectId);
        }

        public IReadOnlyCollection<string> ProjectWhoseCompilationArgumentsChanged()
        {
            return _projectsWhoseCompilationArgumentsChanged;
        }

        public void ClearProjectWhoseCompilationArgumentsChanged()
        {
            _projectsWhoseCompilationArgumentsChanged.Clear();
        }

        public void RemoveCompilationArgumentsFromCache(IEnumerable<string> projectIds)
        {
            foreach (var projectId in projectIds)
            {
                _compilationArguments.Remove(projectId);
            }
        }
    }
}
