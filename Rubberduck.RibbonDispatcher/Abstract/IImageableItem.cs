////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using stdole;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface IImageableItem {
        /// <summary>TODO</summary>
        object  Image     { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        bool    ShowImage { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        bool    ShowLabel { get; set; }

        /// <summary>TODO</summary>
        void SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        void SetImageMso(string imageMso);
    }
}
