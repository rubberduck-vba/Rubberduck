using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserFormCommand : AddComponentCommandBase
    {
        public AddUserFormCommand(ICodeExplorerAddComponentService addComponentService) 
            : base(addComponentService)
        {}

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VBA;

        public override ComponentType ComponentType => ComponentType.UserForm;
    }
}
