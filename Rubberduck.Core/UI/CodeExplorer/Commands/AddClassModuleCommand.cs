using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddClassModuleCommand : AddComponentCommandBase
    {
        public AddClassModuleCommand(IVBE vbe) : base(vbe) { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.All;

        public override ComponentType ComponentType => ComponentType.ClassModule;
    }
}
