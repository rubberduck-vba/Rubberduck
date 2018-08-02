using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.PreProcessing
{
    public class CompilationArgumentsProvider : ICompilationArgumentsProvider
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly IUiDispatcher _uiDispatcher;

        public CompilationArgumentsProvider(IProjectsProvider projectsProvider, IUiDispatcher uiDispatcher)
        {
            _projectsProvider = projectsProvider;
            _uiDispatcher = uiDispatcher;
        }

        public Dictionary<string, short> UserDefinedCompilationArguments(string projectId)
        {
            var project = _projectsProvider.Project(projectId);
            return GetUserDefinedCompilationArguments(project);
        }

        private Dictionary<string, short> GetUserDefinedCompilationArguments(IVBProject project)
        {
            if (project == null)
            {
                return new Dictionary<string, short>();
            }

            // use the TypeLib API to grab the user defined compilation arguments; must be obtained on the main thread.
            var task = _uiDispatcher.StartTask(() => {
                //TODO Push the typelib generation from the project to an ITypeLibProvider taking a projectId and returning the corresponding typeLib.
                using (var typeLib = TypeLibWrapper.FromVBProject(project))
                {
                    return typeLib.ConditionalCompilationArguments;
                }
            });
            return task.Result;
        }
    }
}
