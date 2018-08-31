namespace Rubberduck.VBEditor.SafeComWrappers
{
    /// <summary>
    /// Describes the kind of VB Editor.
    /// </summary>
    /// <remarks>
    /// This is not the same as vbext_ProjectType, exposed by VBProject.Type.
    /// Although the member names are similar, VBProject.Type describes the type of project (either
    /// a hosted project (VBA) or a standalone project (ocx, dll, etc.), not the type of IDE.
    /// </remarks>
    public enum VBEKind
    {
        /// <summary>
        /// Hosted VB editor (Visual Basic for Applications).
        /// </summary>
        Hosted,

        /// <summary>
        /// Standalone VB editor (Visual Basic).
        /// </summary>
        Standalone
    }
}
