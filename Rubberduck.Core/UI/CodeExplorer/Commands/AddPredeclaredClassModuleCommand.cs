using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddPredeclaredClassModuleCommand : CommandBase
    {
        private readonly AddComponentCommand _addComponentCommand;

        public AddPredeclaredClassModuleCommand(AddComponentCommand addComponentCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _addComponentCommand = addComponentCommand;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _addComponentCommand.CanAddComponent(parameter as CodeExplorerItemViewModel, new []{ProjectType.HostProject, ProjectType.StandAlone, ProjectType.StandardExe});
        }

        protected override void OnExecute(object parameter)
        {
            string moduleText = CreatePreclaredClassModule();
            _addComponentCommand.AddComponent(parameter as CodeExplorerItemViewModel, moduleText);
        }

        private string CreatePreclaredClassModule()
        {
            //module text intentionally omits a VB_Name attribute so that the name is automatically assigned by the VBE.
            string moduleText = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = ""Rubberduck"",""Predeclared Class Module""
Option Explicit
";
            return moduleText;
        }
    }
}
