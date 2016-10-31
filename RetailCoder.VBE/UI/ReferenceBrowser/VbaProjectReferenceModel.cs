using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class VbaProjectReferenceModel : AbstractReferenceModel
    {
        private readonly IReference _reference;

        public VbaProjectReferenceModel(IReference reference)
        {
            _reference = reference;
        }

        public static VbaProjectReferenceModel LoadVbaProjectReference(string filePath)
        {
            // TODO load the project from disk.  Get a reference to it from
            throw new NotImplementedException();
        }

        public override string FilePath { get { return _reference.FullPath; } }

        public override string Name { get { return _reference.Name; } }

        public override short MajorVersion { get { return (short) _reference.Major; } }

        public override short MinorVersion { get { return (short) _reference.Minor; } }

        public override Guid Guid { get { return Guid.Parse(_reference.Guid); } }
    }
}