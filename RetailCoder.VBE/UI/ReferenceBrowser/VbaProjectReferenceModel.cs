using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class VbaProjectReferenceModel : VbaReferenceModel
    {
        private readonly VBProject _project;

        internal VbaProjectReferenceModel(VBProject project)
        {
            _project = project;
        }

        public override string FilePath { get { return _project.FileName; } }

        public override string Name { get { return _project.Name; } }

        public override short MajorVersion { get { return 1; } }

        public override short MinorVersion { get { return 0; } }

        public override Guid Guid { get { return Guid.Empty; } }
    }
}