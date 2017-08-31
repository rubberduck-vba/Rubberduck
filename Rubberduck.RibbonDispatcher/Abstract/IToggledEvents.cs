////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("4BD0C027-BD10-4942-B9FA-96A29AB07FE8")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IToggledEvents {
        /// <summary>TODO</summary>
        void Toggled(object sender, IToggledEventArgs e);
    }
}
