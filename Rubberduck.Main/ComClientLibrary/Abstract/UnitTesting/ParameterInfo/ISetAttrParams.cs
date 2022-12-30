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
        Guid(RubberduckGuid.ParamsSetAttrGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always),
    ]
    public interface ISetAttrParams
    {
        /// <summary>
        /// Gets the name of the 'PathName' parameter.
        /// </summary>
        [DispId(1)]
        [Description("Gets the name of the 'PathName' parameter.")]
        string PathName { get; }

        /// <summary>
        /// Gets the name of the 'Attributes' parameter.
        /// </summary>
        [DispId(2)]
        [Description("Gets the name of the 'Attributes' parameter.")]
        string Attributes { get; }
    }
}
