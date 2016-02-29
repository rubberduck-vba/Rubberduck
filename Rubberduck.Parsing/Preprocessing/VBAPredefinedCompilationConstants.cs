using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class VBAPredefinedCompilationConstants
    {
        private readonly string _vbVersion;

        public VBAPredefinedCompilationConstants(string vbVersion)
        {
            _vbVersion = vbVersion;
        }

        public const string VBA6_NAME = "VBA6";
        public const string VBA7_NAME = "VBA7";
        public const string WIN64_NAME = "Win64";
        public const string WIN32_NAME = "Win32";
        public const string WIN16_NAME = "Win16";
        public const string MAC_NAME = "Mac";

        public bool VBA6
        {
            get
            {
                return _vbVersion.StartsWith("6.");
            }
        }

        public bool VBA7
        {
            get
            {
                return _vbVersion.StartsWith("7.");
            }
        }

        public bool Win64
        {
            get
            {
                return Environment.Is64BitOperatingSystem;
            }
        }

        public bool Win32
        {
            get
            {
                return true;
            }
        }

        public bool Win16
        {
            get
            {
                return false;
            }
        }

        public bool Mac
        {
            get
            {
                return false;
            }
        }
    }
}
