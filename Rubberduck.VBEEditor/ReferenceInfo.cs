using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    public readonly struct ReferenceInfo
    {
        public ReferenceInfo(string name, string fullPath, int major, int minor)
        {
            Name = name;
            FullPath = fullPath;
            Major = major;
            Minor = minor;
        }

        public ReferenceInfo(IReference reference)
        :this(reference.Name, 
            reference.FullPath, 
            reference.Major, 
            reference.Minor)
        {}

        public string Name { get; }
        public string FullPath { get; }
        public int Major { get; }
        public int Minor { get; }

        public override int GetHashCode()
        {
            return HashCode.Compute(Name ?? string.Empty, FullPath ?? string.Empty, Major, Minor);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ReferenceInfo other))
            {
                return false;
            }

            return  other.Name == Name
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