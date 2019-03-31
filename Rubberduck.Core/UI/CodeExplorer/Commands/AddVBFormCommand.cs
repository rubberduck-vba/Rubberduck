using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class AddVBFormCommand : AddComponentCommandBase
    {
        public AddVBFormCommand(IVBE vbe) : base(vbe) { }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.VB6;

        public override ComponentType ComponentType => ComponentType.VBForm;
    }
}
