namespace Rubberduck.VBEditor.VbeRuntime.Settings
{
    public interface IVbeSettings
    {
        DllVersion Version { get; }
        bool CompileOnDemand { get; set; }
        bool BackGroundCompile { get; set; }
    }
}