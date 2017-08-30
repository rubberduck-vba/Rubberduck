using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification = "Necessary for COM Interop.")]
    [ComVisible(true)]
    [Guid("B65A7D8F-B46A-45F7-A628-9CB4B84F7EEB")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISelectionMadeEventArgs {
        /// <summary>TODO</summary>
        string ItemId { get; }
    }
}
