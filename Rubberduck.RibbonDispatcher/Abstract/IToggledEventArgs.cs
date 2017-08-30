using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification = "Necessary for COM Interop.")]
    [ComVisible(true)]
    [Guid("A4BDC07D-5A0C-4B98-BA46-46374ABDAF5C")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IToggledEventArgs {
        /// <summary>TODO</summary>
        bool IsPressed { get; }
    }
}
