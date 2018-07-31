using System;
using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.PreProcessing
{
    public class CompilationArgumentsProvider : ICompilationArgumentsProvider
    {
        private readonly IProjectsProvider _projectsProvider;

        public CompilationArgumentsProvider(IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
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
            var providerInst = UiContextProvider.Instance();
            var task = (new UiDispatcher(providerInst)).StartTask(() => {
                using (var typeLib = TypeLibWrapper.FromVBProject(project))
                {
                    return typeLib.ConditionalCompilationArguments;
                }
            });
            return task.Result;
        }
    }
}
