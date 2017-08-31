////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface IDropDownItem {
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        string SelectedItemId      { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.OnActionDropDown)]
        void OnActionDropDown(string itemId);
    }
}
