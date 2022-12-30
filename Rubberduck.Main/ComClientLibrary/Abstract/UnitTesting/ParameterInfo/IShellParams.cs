using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.

namespace Rubberduck.UnitTesting
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.ParamsShellGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always),
    ]
    public interface IShellParams
    {
        /// <summary>
        /// Gets the name of the 'PathName' parameter.
        /// </summary>
        [DispId(1)]
        [Description("Gets the name of the 'PathName' parameter.")]
        string PathName { get; }

        /// <summary>
        /// Gets the name of the 'WindowStyle' optional parameter.
        /// </summary>
        [DispId(2)]
        [Description("Gets the name of the 'WindowStyle' parameter.")]
        string WindowStyle { get; }
    }
}
