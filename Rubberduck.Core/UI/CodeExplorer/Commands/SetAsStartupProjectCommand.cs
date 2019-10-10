using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
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

        public SetAsStartupProjectCommand(
            IVBE vbe, 
            RubberduckParserState parserState, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _parserState = parserState;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            try
            {
                if (!(parameter is CodeExplorerProjectViewModel node) ||
                    !(node.Declaration?.Project is IVBProject project) ||
                    !ProjectTypes.VB6.Contains(project.Type) ||
                    _vbe.ProjectsCount <= 1)
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
            if (!CanExecute(parameter) ||
                !(parameter is CodeExplorerProjectViewModel node))
            {
                return;
            }

            try
            {
                var project = node.Declaration.Project;

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
