using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserControlCommand : AddComponentCommandBase
    {
        public AddUserControlCommand(ICodeExplorerAddComponentService addComponentService) 
            : base(addComponentService)
        {}

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VB6;

        public override ComponentType ComponentType => ComponentType.UserControl;
    }
}
