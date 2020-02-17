using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting.ComCommands
{
    /// <summary>
    /// A command that adds a new test module to the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class AddTestModuleCommand : ComCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ITestCodeGenerator _codeGenerator;
        private readonly IProjectsProvider _projectsProvider;

        public AddTestModuleCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            ITestCodeGenerator codeGenerator,
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider)
            : base(vbeEvents)
        {
            Vbe = vbe;
            _state = state;
            _codeGenerator = codeGenerator;
            _projectsProvider = projectsProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        protected IVBE Vbe { get; }

        private IVBProject GetProject()
        {
            //No using because the wrapper gets returned potentially. 
            var activeProject = Vbe.ActiveVBProject;
            if (!activeProject.IsWrappingNullReference)
            {
                return activeProject;
            }
            activeProject.Dispose();
            
            using (var projects = Vbe.VBProjects)
            {
                return projects.Count == 1
                    ? projects[1] // because VBA-Side indexing
                    : null;
            }
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            bool canExecute;
            using (var project = GetProject())
            {
                canExecute = project != null && !project.IsWrappingNullReference && CanExecuteCode(project);
            }

            return canExecute;
        }
        
        private bool CanExecuteCode(IVBProject project)
        {
            return project.Protection == ProjectProtection.Unprotected;
        }

        protected override void OnExecute(object parameter)
        {
            var parameterIsModuleDeclaration = parameter is ProceduralModuleDeclaration || parameter is ClassModuleDeclaration;

            switch(parameter)
            {
                case IVBProject project:
                    _codeGenerator.AddTestModuleToProject(project);
                    break;
                case Declaration declaration when parameterIsModuleDeclaration:
                    var declarationProject = _projectsProvider.Project(declaration.ProjectId);
                    _codeGenerator.AddTestModuleToProject(declarationProject, declaration);
                    break;
                default:
                    using (var project = GetProject())
                    {
                        _codeGenerator.AddTestModuleToProject(project, null);
                    }
                    break;
            }

            _state.OnParseRequested(this);
        }
    }
}