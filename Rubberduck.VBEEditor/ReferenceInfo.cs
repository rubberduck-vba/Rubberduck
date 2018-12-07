using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    public readonly struct ReferenceInfo
    {
        public ReferenceInfo(Guid guid, string name, string fullPath, int major, int minor)
        {
            Guid = guid;
            Name = name;
            FullPath = fullPath;
            Major = major;
            Minor = minor;
        }

        public ReferenceInfo(IReference reference)
        :this(string.IsNullOrEmpty(reference.Guid) ? Guid.Empty : Guid.Parse(reference.Guid),
            reference.Name, 
            reference.FullPath, 
            reference.Major, 
            reference.Minor)
        {}

        public Guid Guid { get; }
        public string Name { get; }
        public string FullPath { get; }
        public int Major { get; }
        public int Minor { get; }

        public override int GetHashCode()
        {
            return HashCode.Compute(Guid, Name ?? string.Empty, FullPath ?? string.Empty, Major, Minor);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ReferenceInfo other))
            {
                return false;
            }

            return Guid.Equals(other.Guid)
                      && other.Name == Name
                      && other.FullPath == FullPath
                      && other.Major == Major
                      && other.Minor == Minor; 
        }

        public static bool operator ==(ReferenceInfo a, ReferenceInfo b)
        {
            return a.Equals(b);
        }

        public static bool operator !=(ReferenceInfo a, ReferenceInfo b)
        {
            return !a.Equals(b);
        }
    }
}