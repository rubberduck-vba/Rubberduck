using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserFormCommand : AddComponentCommandBase
    {
        public AddUserFormCommand(
            ICodeExplorerAddComponentService addComponentService, 
            IVbeEvents vbeEvents,
            IProjectsProvider projectsProvider) 
            : base(addComponentService, vbeEvents, projectsProvider)
        { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VBA;

        public override ComponentType ComponentType => ComponentType.UserForm;
    }
}
