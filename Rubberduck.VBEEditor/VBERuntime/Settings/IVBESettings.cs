namespace Rubberduck.VBEditor.VBERuntime.Settings
{
    public interface IVBESettings
    {
        DllVersion Version { get; }
        bool CompileOnDemand { get; set; }
        bool BackGroundCompile { get; set; }
    }
}