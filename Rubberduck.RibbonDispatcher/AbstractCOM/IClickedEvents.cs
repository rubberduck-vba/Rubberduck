////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface IClickedEvents {
        /// <summary>TODO</summary>
        void Clicked(object sender, EventArgs e);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(RubberduckGuid.IClickedComEvents)]
    public interface IClickedComEvents {
        /// <summary>TODO</summary>
        [DispId(1)]
        void ComClicked(string ControlId);
    }
}
