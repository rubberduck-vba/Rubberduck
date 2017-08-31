////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("3AD5B841-BA7F-4CFA-9A60-8124B802BF46")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISelectionMadeEvents {
        /// <summary>TODO</summary>
        void SelectionMade(object sender, ISelectionMadeEventArgs e);
    }
}
