using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserDocumentCommand : AddComponentCommandBase
    {
        private static readonly ProjectType[] Types = { ProjectType.ActiveXExe, ProjectType.ActiveXDll };

        public AddUserDocumentCommand(IVBE vbe) : base(vbe) { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => Types;

        public override ComponentType ComponentType => ComponentType.DocObject;
    }
}
