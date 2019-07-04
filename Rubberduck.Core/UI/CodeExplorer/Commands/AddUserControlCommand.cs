using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserControlCommand : AddComponentCommandBase
    {
        public AddUserControlCommand(
            ICodeExplorerAddComponentService addComponentService, IVbeEvents vbeEvents) 
            : base(addComponentService, vbeEvents)
        { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VB6;

        public override ComponentType ComponentType => ComponentType.UserControl;
    }
}
