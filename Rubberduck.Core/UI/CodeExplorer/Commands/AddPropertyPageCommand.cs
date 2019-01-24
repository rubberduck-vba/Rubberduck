using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddPropertyPageCommand : AddComponentCommandBase
    {
        public AddPropertyPageCommand(IVBE vbe) : base(vbe) { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VB6;

        public override ComponentType ComponentType => ComponentType.PropPage;
    }
}
