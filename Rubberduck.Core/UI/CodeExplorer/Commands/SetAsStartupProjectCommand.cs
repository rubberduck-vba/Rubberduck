using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class SetAsStartupProjectCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _parserState;

        public SetAsStartupProjectCommand(IVBE vbe, RubberduckParserState parserState)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _parserState = parserState;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            try
            {
                if (_vbe.ProjectsCount <= 1)
                {
                    return false;
                }

                var project = GetDeclaration(parameter as CodeExplorerItemViewModel)?.Project;

                using (var vbProjects = _vbe.VBProjects)
                {
                    return project != null && !project.Equals(vbProjects.StartProject);
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
            try
            {
                if (_vbe.ProjectsCount <= 1)
                {
                    return;                    
                }

                var project = GetDeclaration(parameter as CodeExplorerItemViewModel)?.Project;

                using (var vbProjects = _vbe.VBProjects)
                {
                    if (!vbProjects.StartProject.Equals(project))
                    {
                        vbProjects.StartProject = project;
                        _parserState.OnParseRequested(this);
                    }
                }
            }
            catch (COMException exception)
            {
                Logger.Error(exception);
            }
        }

        private Declaration GetDeclaration(CodeExplorerItemViewModel node)
        {
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return (node as ICodeExplorerDeclarationViewModel)?.Declaration;
        }
    }
}
