using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("09B49B8B-145A-435D-BE62-17B605D3931A")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IClickedEvents {
        /// <summary>TODO</summary>
        void Clicked(object sender, EventArgs e);
    }
}
