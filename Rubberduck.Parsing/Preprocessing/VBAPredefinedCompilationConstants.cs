using System;
using System.Collections.Generic;

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

        //For some reason, the predefined compilation arguments in VBA are of type Integer
        //with the value 0 for False and 1 for True, which does not correspond to the value
        //a conversion of True from Boolean to Integer would yield. (-1)
        public short VBA7 => (short)(_vbVersion >= 7 ? 1 : 0);
        public short VBA6 => (short)(_vbVersion >= 6 ? 1 : 0);
        public short Win64 = (short)(IntPtr.Size >= 8 ? 1 : 0);
        public short Win32 = 1;
        public short Win16 = 0;
        public short Mac = 0;

        public IDictionary<string, short> AllPredefinedConstants => new Dictionary<string, short>
        {
            {VBA6_NAME, VBA6},
            {VBA7_NAME, VBA7},
            {WIN64_NAME, Win64},
            {WIN32_NAME, Win32},
            {WIN16_NAME, Win16},
            {MAC_NAME, Mac}
        };
    }
}
