////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface IDropDownItem {
        /// <summary>TODO</summary>
        string  SelectedItemId      { get; set; }
        /// <summary>TODO</summary>
        void    OnActionDropDown(string itemId);
    }
}
namespace Rubberduck.RibbonDispatcher.AbstractCOM {
}
