using System;

namespace Rubberduck.Common
{
    public struct WindowsVersion : IComparable<WindowsVersion>
    {
        public static WindowsVersion Windows10 = new WindowsVersion(10, 0, 10240);
        public static WindowsVersion Windows81 = new WindowsVersion(6, 3, 9200);
        public static WindowsVersion Windows8 = new WindowsVersion(6, 2, 9200);
        public static WindowsVersion Windows7_SP1 = new WindowsVersion(6, 1, 7601);
        public static WindowsVersion WindowsVista_SP2 = new WindowsVersion(6, 0, 6002);

        public WindowsVersion(int major, int minor, int build)
        {
            Major = major;
            Minor = minor;
            Build = build;
        }

        public int Major { get; }
        public int Minor { get; }
        public int Build { get; }

        public int CompareTo(WindowsVersion other)
        {
            var majorComparison = Major.CompareTo(other.Major);
            if (majorComparison != 0)
            {
                return majorComparison;
            }

            var minorComparison = Minor.CompareTo(other.Minor);

            return minorComparison != 0
                ? minorComparison
                : Build.CompareTo(other.Build);
        }

        public static bool operator ==(WindowsVersion os1, WindowsVersion os2)
        {
            return os1.CompareTo(os2) == 0;
        }

        public static bool operator !=(WindowsVersion os1, WindowsVersion os2)
        {
            return os1.CompareTo(os2) != 0;
        }

        public static bool operator <(WindowsVersion os1, WindowsVersion os2)
        {
            return os1.CompareTo(os2) < 0;
        }

        public static bool operator >(WindowsVersion os1, WindowsVersion os2)
        {
            return os1.CompareTo(os2) > 0;
        }

        public static bool operator <=(WindowsVersion os1, WindowsVersion os2)
        {
            return os1.CompareTo(os2) <= 0;
        }

        public static bool operator >=(WindowsVersion os1, WindowsVersion os2)
        {
            return os1.CompareTo(os2) >= 0;
        }
    }
}
