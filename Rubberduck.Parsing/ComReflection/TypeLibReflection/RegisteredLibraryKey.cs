using System;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    public struct RegisteredLibraryKey
    {
        public Guid Guid { get; }
        public int Major { get; }
        public int Minor { get; }

        public RegisteredLibraryKey(Guid guid, int major, int minor)
        {
            Guid = guid;
            Major = major;
            Minor = minor;
        }
    }
}