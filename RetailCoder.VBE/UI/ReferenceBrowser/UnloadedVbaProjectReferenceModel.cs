using System;

namespace Rubberduck.UI.ReferenceBrowser
{
    internal class UnloadedVbaProjectReferenceModel : AbstractReferenceModel
    {
        private readonly string _filePath;

        public UnloadedVbaProjectReferenceModel(string filePath)
        {
            _filePath = filePath;
        }

        public override string FilePath
        {
            get { return _filePath; }
        }

        public override string Name
        {
            get { return _filePath; }
        }

        public override short MajorVersion
        {
            get { return 0; }
        }

        public override short MinorVersion
        {
            get { return 0; }
        }

        public override Guid Guid
        {
            get { return Guid.Empty; }
        }
    }
}