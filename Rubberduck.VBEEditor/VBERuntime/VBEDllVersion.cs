using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.VbeRuntime
{
    public enum DllVersion
    {
        Unknown,
        Vb98,
        Vbe6,
        Vbe7
    }

    public static class VbeDllVersion
    {
        public static DllVersion GetCurrentVersion(IVBE vbe)
        {
            try
            {
                switch (int.Parse(vbe.Version.Split('.')[0]))
                {
                    case 6:
                        return vbe.Kind == VBEKind.Standalone ? DllVersion.Vb98 : DllVersion.Vbe6;
                    case 7:
                        return DllVersion.Vbe7;
                    default:
                        return DllVersion.Unknown;
                }
            }
            catch
            {
                return DllVersion.Unknown;
            }
        }
    }
}
