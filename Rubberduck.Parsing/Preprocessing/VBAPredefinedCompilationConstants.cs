using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPredefinedCompilationConstants
    {
        private readonly double _vbVersion;

        public VBAPredefinedCompilationConstants(double vbVersion)
        {
            _vbVersion = vbVersion;
        }

        public const string VBA6_NAME = "VBA6";
        public const string VBA7_NAME = "VBA7";
        public const string WIN64_NAME = "Win64";
        public const string WIN32_NAME = "Win32";
        public const string WIN16_NAME = "Win16";
        public const string MAC_NAME = "Mac";

        public bool VBA7 => _vbVersion < 8;

        public bool VBA6 => _vbVersion < 7;

        public bool Win64 => Environment.Is64BitOperatingSystem;

        public bool Win32 => true;

        public bool Win16 => false;

        public bool Mac => false;
    }
}
