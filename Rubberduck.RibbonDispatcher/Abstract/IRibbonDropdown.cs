////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract
{

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("7660882A-351B-4518-AFD3-8CA1E3EFE9D8")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonDropDown : IRibbonCommon
    {
        /// <summary>TODO</summary>
        string SelectedItemId  { get; set; }

        /// <summary>TODO</summary>
        void   OnAction(string itemId);
    }
}
