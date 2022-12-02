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
        Guid(RubberduckGuid.ParamsSaveSettingGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always),
    ]
    public interface ISaveSettingParams
    {
        /// <summary>
        /// Gets the name of the 'AppName' parameter.
        /// </summary>
        [DispId(1)]
        [Description("Gets the name of the 'AppName' parameter.")]
        string AppName { get; }

        /// <summary>
        /// Gets the name of the 'Section' parameter.
        /// </summary>
        [DispId(2)]
        [Description("Gets the name of the 'Section' parameter.")]
        string Section { get; }

        /// <summary>
        /// Gets the name of the 'Key' parameter.
        /// </summary>
        [DispId(3)]
        [Description("Gets the name of the 'Key' parameter.")]
        string Key { get; }

        /// <summary>
        /// Gets the name of the 'Setting' parameter.
        /// </summary>
        [DispId(4)]
        [Description("Gets the name of the 'Setting' parameter.")]
        string Setting { get; }
    }
}
