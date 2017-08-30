using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IControlChangedEventArgs {
        /// <summary>TODO</summary>
        string ControlId { get; }
    }
}
