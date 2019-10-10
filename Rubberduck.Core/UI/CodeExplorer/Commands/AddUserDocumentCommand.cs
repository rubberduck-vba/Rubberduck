using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserDocumentCommand : AddComponentCommandBase
    {
        private static readonly ProjectType[] Types = { ProjectType.ActiveXExe, ProjectType.ActiveXDll };

        public AddUserDocumentCommand(
            ICodeExplorerAddComponentService addComponentService, IVbeEvents vbeEvents) 
            : base(addComponentService, vbeEvents)
        { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.DocObject;
    }
}
