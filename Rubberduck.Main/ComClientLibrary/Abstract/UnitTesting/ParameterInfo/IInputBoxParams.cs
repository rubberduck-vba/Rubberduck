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
    Guid(RubberduckGuid.ParamsInputBoxGuid),
    InterfaceType(ComInterfaceType.InterfaceIsDual),
    EditorBrowsable(EditorBrowsableState.Always),   
    ]
    public interface IInputBoxParams
    {
        /// <summary>
        /// Gets the name of the 'Prompt' parameter.
        /// </summary>
        [DispId(1)]
        [Description("Gets the name of the 'Prompt' parameter.")]
        string Prompt { get; }

        /// <summary>
        /// Gets the name of the 'Title' optional parameter.
        /// </summary>
        [DispId(2)]
        [Description("Gets the name of the 'Title' optional parameter.")]
        string Title { get; }

        /// <summary>
        /// Gets the name of the 'Default' optional parameter.
        /// </summary>
        [DispId(3)]
        [Description("Gets the name of the 'Default' optional parameter.")]
        string Default { get; }

        /// <summary>
        /// Gets the name of the 'XPos' optional parameter.
        /// </summary>
        [DispId(4)]
        [Description("Gets the name of the 'XPos' optional parameter.")]
        string XPos { get; }

        /// <summary>
        /// Gets the name of the 'YPos' optional parameter.
        /// </summary>
        [DispId(5)]
        [Description("Gets the name of the 'YPos' optional parameter.")]
        string YPos { get; }

        /// <summary>
        /// Gets the name of the 'HelpFile' optional parameter.
        /// </summary>
        [DispId(6)]
        [Description("Gets the name of the 'HelpFile' optional parameter.")]
        string HelpFile { get; }

        /// <summary>
        /// Gets the name of the 'Context' optional parameter.
        /// </summary>
        [DispId(7)]
        [Description("Gets the name of the 'Context' optional parameter.")]
        string Context { get; }
    }
}
