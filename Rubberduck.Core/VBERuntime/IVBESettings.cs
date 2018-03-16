namespace Rubberduck.VBERuntime
{
    public interface IVBESettings
    {
        VBESettings.DllVersion Version { get; }
        bool CompileOnDemand { get; set; }
        bool BackGroundCompile { get; set; }
    }
}