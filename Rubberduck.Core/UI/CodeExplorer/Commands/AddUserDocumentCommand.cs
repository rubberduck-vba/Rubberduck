using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserDocumentCommand : AddComponentCommandBase
    {
        private static readonly ProjectType[] Types = { ProjectType.ActiveXExe, ProjectType.ActiveXDll };

        public AddUserDocumentCommand(ICodeExplorerAddComponentService addComponentService) 
            : base(addComponentService)
        {}

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.DocObject;
    }
}
