using System;

namespace Rubberduck.VBEditor.VbeRuntime
{
    public interface IVbeNativeApi
    {
        string DllName { get; }
        int DoEvents();
    }
}
