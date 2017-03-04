namespace Rubberduck.VBEditor.SafeComWrappers
{
    /// <summary>
    /// Values compatible with <see cref="Microsoft.Vbe.Interop.vbext_VBAMode"/> enum values.
    /// </summary>
    public enum EnvironmentMode
    {
        Run = 0,
        Break = 1,
        Design = 2,
    }
}