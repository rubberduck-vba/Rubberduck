////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(RubberduckGuid.IToggledEvents)]
    public interface IToggledEvents {
        /// <summary>TODO</summary>
        [DispId(1)]
        void Toggled(object sender, IToggledEventArgs e);
    }
}
