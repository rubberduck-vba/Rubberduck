using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting.ComCommands;
using Rubberduck.UnitTesting.CodeGeneration;
using Rubberduck.VBEditor.Events;
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

        public AddTestComponentCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            ITestCodeGenerator codeGenerator, 
            IVbeEvents vbeEvents)
            : base(vbe, state, codeGenerator, vbeEvents)
        {
            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (parameter == null)
            {
                return false;
            }

            Declaration declaration;
            if (ApplicableNodes.Contains(parameter.GetType()) &&
                parameter is CodeExplorerItemViewModel node)
            {
                declaration = node.Declaration;
            }
            else if (parameter is Declaration d)
            {
                declaration = d;
            }
            else
            {
                return false;
            }

            try
            {
                return declaration?.Project != null || Vbe.ProjectsCount != 1;
            }
            catch (COMException)
            {
                return false;
            }
        }
    }
}
