using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class SetAsStartupProjectCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes = { typeof(CodeExplorerProjectViewModel) };

        private readonly IVBE _vbe;
        private readonly RubberduckParserState _parserState;
        private readonly IProjectsProvider _projectsProvider;

        public SetAsStartupProjectCommand(
            IVBE vbe, 
            RubberduckParserState parserState, 
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _parserState = parserState;
            _projectsProvider = projectsProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            try
            {
                if (!(parameter is CodeExplorerProjectViewModel node) 
                    || node.Declaration == null
                    || _vbe.ProjectsCount <= 1)
                {
                    return false;
                }

                var project = _projectsProvider.Project(node.Declaration.ProjectId);
                if (project == null 
                    || !ProjectTypes.VB6.Contains(project.Type))
                {
                    return false;
                }

                using (var vbProjects = _vbe.VBProjects)
                {
                    return !project.Equals(vbProjects.StartProject);
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception);
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter) 
                || !(parameter is CodeExplorerProjectViewModel node)
                || node.Declaration == null)
            {
                return;
            }

            var project = _projectsProvider.Project(node.Declaration.ProjectId);
            if (project == null)
            {
                return;
            }

            try
            {
                using (var vbProjects = _vbe.VBProjects)
                {
                    vbProjects.StartProject = project;
                    _parserState.OnParseRequested(this);
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception);
            }
        }
    }
}
