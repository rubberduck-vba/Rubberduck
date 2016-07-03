using System;

namespace Rubberduck.UI.ReferenceBrowser
{
    public abstract class VbaReferenceModel
    {
        public abstract string FilePath { get; }

        public abstract string Name { get; }

        public abstract short MajorVersion { get; }

        public abstract short MinorVersion { get; }

        public abstract Guid Guid { get; }
    }
}