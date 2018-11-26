namespace Rubberduck.VBEditor.SafeComWrappers
{
    /// <summary>
    /// Values compatible with <see cref="Microsoft.Vbe.Interop.vbext_ProjectType"/> enum values.
    /// </summary>
    public enum ProjectType
    {        
        StandardExe = 0,
        ActiveXExe = 1,
        ActiveXDll = 2,
        ActiveXControl = 3,
        HostProject = 100,
        StandAlone = 101
    }
}