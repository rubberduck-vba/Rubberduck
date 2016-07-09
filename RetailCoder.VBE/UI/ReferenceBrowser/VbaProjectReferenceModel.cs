using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class VbaProjectReferenceModel : VbaReferenceModel
    {
        private readonly Reference _reference;

        public VbaProjectReferenceModel(Reference reference)
        {
            _reference = reference;
        }

        public override string FilePath { get { return _reference.FullPath; } }

        public override string Name { get { return _reference.Name; } }

        public override short MajorVersion { get { return (short) _reference.Major; } }

        public override short MinorVersion { get { return (short) _reference.Minor; } }

        public override Guid Guid { get { return Guid.Parse(_reference.Guid); } }
    }
}