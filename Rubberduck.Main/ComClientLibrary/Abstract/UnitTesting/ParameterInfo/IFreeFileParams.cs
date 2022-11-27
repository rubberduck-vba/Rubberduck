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
        Guid(RubberduckGuid.ParamsFreeFileGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always),
    ]
    public interface IFreeFileParams
    {
        /// <summary>
        /// Gets the name of the 'RangeNumber' parameter.
        /// </summary>
        [DispId(1)]
        [Description("Gets the name of the 'RangeNumber' parameter.")]
        string RangeNumber { get; }
    }
}
