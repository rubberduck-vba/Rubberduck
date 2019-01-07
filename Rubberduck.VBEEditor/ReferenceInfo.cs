using System;
using System.Runtime.Serialization;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable NonReadonlyMemberInGetHashCode
namespace Rubberduck.VBEditor
{
    [DataContract]
    public struct ReferenceInfo
    {
        public static ReferenceInfo Empty => new ReferenceInfo(Guid.Empty, string.Empty, string.Empty, 0, 0);

        public ReferenceInfo(Guid guid, string name, string fullPath, int major, int minor)
        {
            Guid = guid;
            Name = name;
            FullPath = fullPath;
            Major = major;
            Minor = minor;
        }

        public ReferenceInfo(IReference reference)
        {
            Guid = Guid.TryParse(reference.Guid, out var guid) ? guid : Guid.Empty;
            Name = reference.Name;
            FullPath = reference.FullPath;
            Major = reference.Major;
            Minor = reference.Minor;
        }

        [DataMember(IsRequired = true)]
        public Guid Guid { get; private set; }

        [DataMember(IsRequired = true)]
        public string Name { get; private set; }

        [DataMember(IsRequired = true)]
        public string FullPath { get; private set; }

        [DataMember(IsRequired = true)]
        public int Major { get; private set; }

        [DataMember(IsRequired = true)]
        public int Minor { get; private set; }

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