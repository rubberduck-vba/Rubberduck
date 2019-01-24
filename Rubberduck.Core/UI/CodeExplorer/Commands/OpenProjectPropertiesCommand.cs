using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class OpenProjectPropertiesCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly IVBE _vbe;

        public OpenProjectPropertiesCommand(IVBE vbe)
        {
            _vbe = vbe;
        }

        public sealed override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (!base.EvaluateCanExecute(parameter) ||
                !(parameter is CodeExplorerItemViewModel node))
            {
                return false;
            }

            try
            {
                return node.Declaration != null || _vbe.ProjectsCount == 1;
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            const int openProjectPropertiesId = 2578;

            using (var commandBars = _vbe.CommandBars)
            {
                using (var command = commandBars.FindControl(openProjectPropertiesId))
                {
                    if (_vbe.ProjectsCount == 1)
                    {
                        command.Execute();
                        return;
                    }

                    if (!(parameter is CodeExplorerItemViewModel node))
                    {
                        return;
                    }

                    var nodeProject = node.Declaration?.Project;
                    if (nodeProject == null)
                    {
                        return; //The project declaration has been disposed, i.e. the project has been removed already.
                    }

                    try
                    {
                        _vbe.ActiveVBProject = nodeProject;
                    }
                    catch (COMException)
                    {
                        return; // the project was probably removed from the VBE, but not from the CE
                    }

                    command.Execute();
                }
            }
        }
    }
}
