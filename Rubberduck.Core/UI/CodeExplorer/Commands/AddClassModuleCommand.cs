using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddClassModuleCommand : AddComponentCommandBase
    {
        public AddClassModuleCommand(ICodeExplorerAddComponentService addComponentService) 
            : base(addComponentService)
        {}

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.All;

        public override ComponentType ComponentType => ComponentType.ClassModule;
    }
}
