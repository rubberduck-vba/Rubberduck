using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddStdModuleCommand : AddComponentCommandBase
    {
        public AddStdModuleCommand(ICodeExplorerAddComponentService addComponentService) 
            : base(addComponentService)
        {}

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.All;

        public override ComponentType ComponentType => ComponentType.StandardModule;
    }
}
