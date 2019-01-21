using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddTestComponentCommand : AddTestModuleCommand
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerCustomFolderViewModel),
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        public AddTestComponentCommand(IVBE vbe, RubberduckParserState state, IGeneralConfigService configLoader, IMessageBox messageBox, IVBEInteraction interaction) 
            : base(vbe, state, configLoader, messageBox, interaction) { }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (parameter == null || 
                !ApplicableNodes.Contains(parameter.GetType()) ||
                !(parameter is CodeExplorerItemViewModel node))
            {
                return false;
            }

            try
            {
                return node.Declaration?.Project != null || Vbe.ProjectsCount != 1;
            }
            catch (COMException)
            {
                return false;
            }
        }
    }
}
