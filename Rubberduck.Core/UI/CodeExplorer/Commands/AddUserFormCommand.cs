using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserFormCommand : AddComponentCommandBase
    {
        public AddUserFormCommand(IVBE vbe, IVBEEvents vbeEvents) : base(vbe, vbeEvents) { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VBA;

        public override ComponentType ComponentType => ComponentType.UserForm;
    }
}
